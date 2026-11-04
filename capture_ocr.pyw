# -*- coding: utf-8 -*-
"""
CAPTURE OCR -> WORD
Runs in the background. Hold ALT + RIGHT mouse button and drag to draw a
selection rectangle. When the right button is released, the rectangle is
OCRed (Italian + English) and the recognized text opens in a NEW, UNSAVED
Microsoft Word document (no file written to disk), with low-confidence
words shown in red and line breaks preserved.
Quit the program with CTRL + ALT + Q (asks for confirmation).
"""

import ctypes
import os
import queue
import sys
import time
import traceback
from datetime import datetime

# ---- DPI awareness: must run BEFORE importing tkinter, otherwise mouse
# ---- and screenshot coordinates do not match on scaled displays ----
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from pynput import keyboard, mouse
import mss
import mss.tools
from PIL import Image
import pytesseract
from docx import Document
from docx.shared import RGBColor


def _msgbox(text, title="Capture OCR", icon=0x10):
    """Always-on-top message box (0x10 = error, 0x30 = warning, 0x40 = info)."""
    try:
        ctypes.windll.user32.MessageBoxW(0, str(text), title,
                                         icon | 0x1000 | 0x40000)
    except Exception:
        pass


class _GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort), ("Data4", ctypes.c_ubyte * 8)]


def desktop_folder():
    """The user's REAL Desktop (handles OneDrive folder redirection too),
    via SHGetKnownFolderPath(FOLDERID_Desktop)."""
    try:
        guid = _GUID(0xB4BFCC3A, 0xDB2C, 0x424C,
                     (ctypes.c_ubyte * 8)(0xB0, 0x29, 0x7F, 0xE9,
                                          0x9A, 0x87, 0xC6, 0x41))
        p = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(
                ctypes.byref(guid), 0, None, ctypes.byref(p)) == 0:
            path = p.value
            ctypes.windll.ole32.CoTaskMemFree(p)
            if path and os.path.isdir(path):
                return path
    except Exception:
        pass
    return os.path.join(os.path.expanduser("~"), "Desktop")


# ================== CONFIGURATION ==================
APP_VERSION = "1.0.0"       # bump this before building each new release
UPDATE_CHECK = True         # on startup, ask GitHub (api.github.com) whether
                            # a newer release exists and offer its download
                            # page. One HTTPS request, nothing else is sent.
                            # False = the program makes NO network access.
GITHUB_REPO = "astrozullo/capture-ocr"  # owner/repo whose Releases are checked

# Tesseract path (standard Windows installation)
TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Tesseract lookup order:
# 1) embedded INSIDE the executable (the "all-in-one" build);
# 2) a "Tesseract-OCR" folder next to the executable/script (portable);
# 3) the standard installation in C:\Program Files (path above);
# 4) fallback: per-user installation in AppData\Local\Programs.
def _tesseract_candidates():
    cand = []
    if getattr(sys, "frozen", False):          # PyInstaller executable
        inner = getattr(sys, "_MEIPASS", "")    # data extracted from the exe
        if inner:
            cand.append(os.path.join(inner, "Tesseract-OCR",
                                     "tesseract.exe"))
        cand.append(os.path.join(os.path.dirname(sys.executable),
                                 "Tesseract-OCR", "tesseract.exe"))
    else:
        cand.append(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "Tesseract-OCR", "tesseract.exe"))
    return cand

for _c in _tesseract_candidates():
    if os.path.exists(_c):
        TESSERACT_EXE = _c
        break

# fallback: per-user ("install for me only") Tesseract installation
if not os.path.exists(TESSERACT_EXE):
    _ua = os.path.join(os.environ.get("LOCALAPPDATA", ""),
                       "Programs", "Tesseract-OCR", "tesseract.exe")
    if os.path.exists(_ua):
        TESSERACT_EXE = _ua

OCR_LANGUAGES = "ita+eng"  # 'eng' always ships with Tesseract; helps on logos
                           # and English terms. Use "eng" for English only.
# Destination folder for Word files (the user's real Desktop)
OUTPUT_FOLDER = desktop_folder()
OUTPUT_MODE = "both"        # where the recognized text goes:
                            #   "clipboard" = copied to the clipboard only
                            #                 (plain text, no red marking)
                            #   "word"      = opened in Microsoft Word only
                            #                 (doubtful words in red)
                            #   "both"      = clipboard AND Word; if Word is
                            #                 missing, the clipboard still
                            #                 works
OPEN_FILE_WHEN_DONE = True  # open the result in Word as soon as it is ready
OPEN_WITHOUT_FILE = True    # True = the text goes into a NEW, UNSAVED Word
                            # document (zero files on disk: you save it from
                            # Word only if you need it). If Word does not
                            # respond, falls back to the file method.
SAVE_TO_DESKTOP = False     # (file method only) False = the .docx goes into
                            # a TEMPORARY folder and opens from there: nothing
                            # is left on the Desktop; use "Save as" from Word
                            # to keep the text. True = file left on the
                            # Desktop.
TEMP_CLEANUP_DAYS = 2       # on startup, delete temp files older than this
DOUBT_THRESHOLD = 60        # words with OCR confidence below this threshold
                            # (0-100 scale) are written in RED in Word
SHOW_STATUS_WIDGET = False  # True = small translucent draggable status
                            # window; False = background only
MIN_SIZE_PX = 8             # smaller rectangles are ignored
OVERLAY_ALPHA = 0.30        # opacity of the dark veil during selection
# ===================================================

if os.path.exists(TESSERACT_EXE):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE

ERROR_LOG = os.path.join(os.path.expanduser("~"), "capture_ocr_errors.log")

# ---- State shared between the listener threads and the tkinter main loop ----
events = queue.Queue()
alt_down = False
capturing = False

WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
WM_MOUSEMOVE = 0x0200

user32 = ctypes.windll.user32
SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN = 76, 77
SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN = 78, 79


def beep_ok():
    try:
        import winsound
        winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        pass


def beep_err():
    try:
        import winsound
        winsound.MessageBeep(winsound.MB_ICONHAND)
    except Exception:
        pass


# ------------------- Keyboard listener -------------------
ALT_KEYS = {keyboard.Key.alt_l, keyboard.Key.alt_r, keyboard.Key.alt_gr}
current_keys = set()


def on_key_press(key):
    global alt_down
    current_keys.add(key)
    if key in ALT_KEYS:
        alt_down = True
    # CTRL + ALT + Q to quit
    try:
        if (key == keyboard.KeyCode.from_char('q') or
                getattr(key, 'vk', None) == 0x51):
            ctrl = any(k in current_keys for k in
                       (keyboard.Key.ctrl_l, keyboard.Key.ctrl_r))
            if ctrl and alt_down:
                events.put(("quit", None))
    except Exception:
        pass


def on_key_release(key):
    global alt_down
    current_keys.discard(key)
    if key in ALT_KEYS:
        alt_down = False


# ------------------- Mouse listener -------------------
def win32_event_filter(msg, data):
    """Intercepts the right button while ALT is held and SUPPRESSES it,
    so the context menu of the underlying application does not open."""
    global capturing
    if msg == WM_RBUTTONDOWN and alt_down:
        capturing = True
        events.put(("start", (data.pt.x, data.pt.y)))
        mouse_listener.suppress_event()
    elif msg == WM_MOUSEMOVE and capturing:
        events.put(("move", (data.pt.x, data.pt.y)))
    elif msg == WM_RBUTTONUP and capturing:
        capturing = False
        events.put(("end", (data.pt.x, data.pt.y)))
        mouse_listener.suppress_event()
    return True


mouse_listener = mouse.Listener(win32_event_filter=win32_event_filter)


# ------------------- OCR + Word -------------------
def _otsu_threshold(gray):
    """Automatic binarization threshold (Otsu's method) on an L-mode image."""
    hist = gray.histogram()
    total = sum(hist)
    sum_total = sum(i * h for i, h in enumerate(hist))
    sum_bg = 0.0
    weight_bg = 0
    var_max = 0.0
    threshold = 128
    for i in range(256):
        weight_bg += hist[i]
        if weight_bg == 0:
            continue
        weight_fg = total - weight_bg
        if weight_fg == 0:
            break
        sum_bg += i * hist[i]
        mean_bg = sum_bg / weight_bg
        mean_fg = (sum_total - sum_bg) / weight_fg
        var = weight_bg * weight_fg * (mean_bg - mean_fg) ** 2
        if var > var_max:
            var_max = var
            threshold = i
    return threshold


def _variants(img):
    """Three versions of the capture: grayscale, binarized, inverted
    binarized (for light text on dark backgrounds)."""
    from PIL import ImageOps
    base = img.resize((img.width * 3, img.height * 3), Image.LANCZOS)
    gray = ImageOps.autocontrast(base.convert("L"))
    t = _otsu_threshold(gray)
    bw = gray.point(lambda p: 255 if p > t else 0)
    return [gray, bw, ImageOps.invert(bw)]


def _lines_from_data(data):
    """Rebuilds the text lines from the image_to_data dictionary, preserving
    line breaks. Each line is a list of (word, confidence) pairs."""
    lines = {}
    for i in range(len(data["text"])):
        word = data["text"][i].strip()
        if not word:
            continue
        try:
            conf = float(data["conf"][i])
        except (TypeError, ValueError):
            conf = -1.0
        key = (data["block_num"][i], data["par_num"][i],
               data["line_num"][i])
        lines.setdefault(key, []).append((word, conf))
    return [lines[k] for k in sorted(lines)]


def best_ocr(img):
    """Runs the OCR on several variants of the image and with two
    segmentation modes (psm 6 = uniform block, psm 11 = sparse text), then
    keeps the result with the highest mean confidence.
    Returns a list of lines; each line = a list of (word, confidence)."""
    best_lines = []
    best_score = (-1.0, 0)   # (mean confidence, character count)
    for var in _variants(img):
        for psm in (6, 11):
            cfg = f"--oem 3 --psm {psm}"
            try:
                data = pytesseract.image_to_data(
                    var, lang=OCR_LANGUAGES, config=cfg,
                    output_type=pytesseract.Output.DICT)
            except Exception:
                continue
            confs = []
            for c in data["conf"]:
                try:
                    c = float(c)
                except (TypeError, ValueError):
                    continue
                if c >= 0:
                    confs.append(c)
            if not confs:
                continue
            lines = _lines_from_data(data)
            n_char = sum(len(w) for ln in lines for w, _ in ln)
            score = (sum(confs) / len(confs), n_char)
            if n_char and score > best_score:
                best_score = score
                best_lines = lines
    return best_lines


def _copy_to_clipboard(text):
    """Puts plain Unicode text on the Windows clipboard (no dependencies)."""
    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002
    k32 = ctypes.windll.kernel32
    u32 = ctypes.windll.user32
    k32.GlobalAlloc.restype = ctypes.c_void_p
    k32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
    k32.GlobalLock.restype = ctypes.c_void_p
    k32.GlobalLock.argtypes = [ctypes.c_void_p]
    k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    u32.SetClipboardData.restype = ctypes.c_void_p
    u32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
    data = text.replace("\n", "\r\n")
    if not u32.OpenClipboard(None):
        raise RuntimeError("OpenClipboard failed")
    try:
        u32.EmptyClipboard()
        buf = ctypes.create_unicode_buffer(data)
        size = ctypes.sizeof(buf)
        handle = k32.GlobalAlloc(GMEM_MOVEABLE, size)
        if not handle:
            raise RuntimeError("GlobalAlloc failed")
        ptr = k32.GlobalLock(handle)
        ctypes.memmove(ptr, buf, size)
        k32.GlobalUnlock(handle)
        if not u32.SetClipboardData(CF_UNICODETEXT, handle):
            raise RuntimeError("SetClipboardData failed")
    finally:
        u32.CloseClipboard()


def _open_in_word_unsaved(lines):
    """Creates a NEW, UNSAVED document in Word with the recognized text
    (doubtful words in red). No file is written to disk."""
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Add()
    for line in lines:
        for j, (w, conf) in enumerate(line):
            txt = ("" if j == 0 else " ") + w
            r = doc.Content
            start = r.End - 1          # before the final paragraph mark
            ins = doc.Range(start, start)
            ins.Text = txt
            colored = doc.Range(start, start + len(txt))
            if 0 <= conf < DOUBT_THRESHOLD:
                colored.Font.Color = 204         # RGB(204,0,0) = red
            else:
                colored.Font.Color = 0           # black
        r = doc.Content
        end = doc.Range(r.End - 1, r.End - 1)
        end.Text = "\r"                           # line break
    try:
        word.Activate()
    except Exception:
        pass


def ocr_and_deliver(bbox):
    """bbox = (left, top, width, height) in screen coordinates."""
    left, top, w, h = bbox
    with mss.mss() as sct:
        shot = sct.grab({"left": left, "top": top, "width": w, "height": h})
        img = Image.frombytes("RGB", shot.size, shot.rgb)

    lines = best_ocr(img)

    if not lines:
        beep_err()
        _msgbox("No text recognized in the selected rectangle.",
                icon=0x30)
        return None

    delivered_clip = False
    delivered_word = None

    # ---- clipboard delivery ----
    if OUTPUT_MODE in ("clipboard", "both"):
        try:
            text = "\n".join(" ".join(w for w, _ in line) for line in lines)
            _copy_to_clipboard(text)
            delivered_clip = True
        except Exception:
            with open(ERROR_LOG, "a", encoding="utf-8") as f:
                f.write("Clipboard copy failed:\n" +
                        traceback.format_exc() + "\n")

    # ---- Word delivery ----
    if OUTPUT_MODE in ("word", "both"):
        if OPEN_WITHOUT_FILE:
            try:
                _open_in_word_unsaved(lines)
                delivered_word = "(unsaved Word document)"
            except Exception:
                # Word/COM not available: fall back to the file method
                with open(ERROR_LOG, "a", encoding="utf-8") as f:
                    f.write("Falling back to file, COM failed:\n" +
                            traceback.format_exc() + "\n")
        if delivered_word is None:
            try:
                doc = Document()
                for line in lines:
                    par = doc.add_paragraph()
                    for j, (w_, conf) in enumerate(line):
                        run = par.add_run(("" if j == 0 else " ") + w_)
                        if 0 <= conf < DOUBT_THRESHOLD:
                            # doubtful word: in red
                            run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

                if SAVE_TO_DESKTOP:
                    destination = OUTPUT_FOLDER
                else:
                    import tempfile
                    destination = os.path.join(tempfile.gettempdir(),
                                               "CaptureOCR")
                os.makedirs(destination, exist_ok=True)
                name = ("Capture_OCR_" +
                        datetime.now().strftime("%Y-%m-%d_%H-%M-%S") +
                        ".docx")
                path = os.path.join(destination, name)
                doc.save(path)
                if OPEN_FILE_WHEN_DONE:
                    try:
                        os.startfile(path)   # open the .docx with Word
                    except Exception:
                        pass
                delivered_word = path
            except Exception:
                with open(ERROR_LOG, "a", encoding="utf-8") as f:
                    f.write("File method failed:\n" +
                            traceback.format_exc() + "\n")

    if delivered_word or delivered_clip:
        beep_ok()
        return delivered_word or "(text copied to the clipboard)"

    beep_err()
    _msgbox("Could not deliver the recognized text (no output mode "
            "succeeded).\n\nFull details in the file "
            "capture_ocr_errors.log in your user folder.")
    return None


# --------- Selection overlay (tkinter, main thread) ---------
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.overlay = None
        self.canvas = None
        self.rect_id = None
        self.p0 = None
        self.vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
        self.vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
        self.vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
        self.vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
        self.widget = None
        self.lbl = None
        if SHOW_STATUS_WIDGET:
            self.create_widget()
        self.root.after(15, self.poll)

    # ---------- status widget ----------
    def create_widget(self):
        w = tk.Toplevel(self.root)
        w.overrideredirect(True)
        w.attributes("-topmost", True)
        w.attributes("-alpha", 0.82)
        w.configure(bg="#202225")
        self.lbl = tk.Label(w, text="OCR  ready", bg="#202225",
                            fg="#7CE8A0", font=("Segoe UI", 9))
        self.lbl.pack(side="left", padx=(10, 6), pady=4)
        close = tk.Label(w, text="\u2715", bg="#202225", fg="#aaaaaa",
                         font=("Segoe UI", 9), cursor="hand2")
        close.pack(side="right", padx=(0, 8))
        close.bind("<Button-1>", lambda e: events.put(("quit", None)))
        for wid in (w, self.lbl):
            wid.bind("<Button-1>", self._drag_start)
            wid.bind("<B1-Motion>", self._drag_move)
        # initial position: bottom right of the primary screen
        w.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w.geometry(f"+{sw - w.winfo_reqwidth() - 18}"
                   f"+{sh - w.winfo_reqheight() - 64}")
        self.widget = w

    def _drag_start(self, e):
        self._dx = e.x_root - self.widget.winfo_x()
        self._dy = e.y_root - self.widget.winfo_y()

    def _drag_move(self, e):
        self.widget.geometry(f"+{e.x_root - self._dx}+{e.y_root - self._dy}")

    def set_status(self, text, color="#7CE8A0"):
        if self.lbl:
            self.lbl.config(text=text, fg=color)
            self.root.update_idletasks()
            self.root.update()

    def hide_widget(self):
        if self.widget:
            self.widget.withdraw()
            self.root.update_idletasks()
            self.root.update()

    def show_widget(self):
        if self.widget:
            self.widget.deiconify()
            self.widget.attributes("-topmost", True)

    def open_overlay(self, p0):
        # the status widget hides itself during the capture, so it can
        # never end up inside the screenshot
        self.hide_widget()
        self.p0 = p0
        self.overlay = tk.Toplevel(self.root)
        self.overlay.overrideredirect(True)
        self.overlay.geometry(f"{self.vw}x{self.vh}+{self.vx}+{self.vy}")
        self.overlay.attributes("-topmost", True)
        self.overlay.attributes("-alpha", OVERLAY_ALPHA)
        self.overlay.configure(bg="black")
        self.canvas = tk.Canvas(self.overlay, bg="black",
                                highlightthickness=0, cursor="crosshair")
        self.canvas.pack(fill="both", expand=True)
        x, y = p0[0] - self.vx, p0[1] - self.vy
        self.rect_id = self.canvas.create_rectangle(
            x, y, x, y, outline="#ff3333", width=2)

    def update_overlay(self, p):
        if self.overlay and self.rect_id:
            x0, y0 = self.p0[0] - self.vx, self.p0[1] - self.vy
            x1, y1 = p[0] - self.vx, p[1] - self.vy
            self.canvas.coords(self.rect_id, x0, y0, x1, y1)

    def close_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None
            self.canvas = None
            self.rect_id = None
        # make sure the overlay is really gone before the screenshot
        self.root.update_idletasks()
        self.root.update()
        time.sleep(0.12)

    def poll(self):
        try:
            while True:
                kind, payload = events.get_nowait()
                if kind == "start":
                    self.open_overlay(payload)
                elif kind == "move":
                    self.update_overlay(payload)
                elif kind == "end":
                    p0 = self.p0
                    self.close_overlay()
                    outcome = ("ready", "#7CE8A0")
                    if p0 is not None:
                        left = min(p0[0], payload[0])
                        top = min(p0[1], payload[1])
                        w = abs(payload[0] - p0[0])
                        h = abs(payload[1] - p0[1])
                        if w >= MIN_SIZE_PX and h >= MIN_SIZE_PX:
                            try:
                                result = ocr_and_deliver((left, top, w, h))
                                if result:
                                    outcome = ("done " +
                                               datetime.now().strftime("%H:%M"),
                                               "#7CE8A0")
                                else:
                                    outcome = ("no text", "#E8D07C")
                            except Exception as e:
                                outcome = ("ERROR", "#F08080")
                                beep_err()
                                with open(ERROR_LOG, "a",
                                          encoding="utf-8") as f:
                                    f.write(traceback.format_exc() + "\n")
                                _msgbox("Error during OCR/delivery:\n\n"
                                        f"{type(e).__name__}: {e}\n\n"
                                        "Full details in the file "
                                        "capture_ocr_errors.log in your "
                                        "user folder.")
                    self.p0 = None
                    self.show_widget()
                    self.set_status("OCR  " + outcome[0], outcome[1])
                elif kind == "quit":
                    if not getattr(self, "confirm_open", False):
                        self.confirm_open = True
                        r = ctypes.windll.user32.MessageBoxW(
                            0,
                            "Do you want to quit Capture OCR?\n\n"
                            "After quitting, ALT + right button will not "
                            "capture anything until you launch the "
                            "program again.",
                            "Capture OCR - confirm quit",
                            0x4 | 0x20 | 0x1000 | 0x40000)  # Yes/No, question
                        self.confirm_open = False
                        if r == 6:  # IDYES
                            self.root.destroy()
                            return
        except queue.Empty:
            pass
        self.root.after(15, self.poll)

    def run(self):
        self.root.mainloop()


def _cleanup_temp():
    """Deletes temporary .docx files older than TEMP_CLEANUP_DAYS.
    Files still open in Word cannot be deleted: they are skipped."""
    import tempfile
    folder = os.path.join(tempfile.gettempdir(), "CaptureOCR")
    if not os.path.isdir(folder):
        return
    limit = time.time() - TEMP_CLEANUP_DAYS * 86400
    for f in os.listdir(folder):
        p = os.path.join(folder, f)
        try:
            if os.path.isfile(p) and os.path.getmtime(p) < limit:
                os.remove(p)
        except Exception:
            pass


def _parse_version(s):
    """'v1.2.10' -> (1, 2, 10). Non-numeric tails are ignored."""
    parts = []
    for piece in s.strip().lstrip("vV").split("."):
        digits = ""
        for ch in piece:
            if ch.isdigit():
                digits += ch
            else:
                break
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def _check_updates():
    """Asks the GitHub API for the latest release of GITHUB_REPO and, if it
    is newer than APP_VERSION, offers to open its download page. Runs in a
    background thread; silent on any failure (offline, no releases yet,
    API rate limit). Nothing about the user or the captures is sent."""
    try:
        import json
        import urllib.request
        req = urllib.request.Request(
            f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
            headers={"User-Agent": "CaptureOCR-update-check",
                     "Accept": "application/vnd.github+json"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            info = json.load(resp)
        latest = info.get("tag_name", "")
        page = info.get("html_url",
                        f"https://github.com/{GITHUB_REPO}/releases")
        if latest and _parse_version(latest) > _parse_version(APP_VERSION):
            r = ctypes.windll.user32.MessageBoxW(
                0,
                f"A new version of Capture OCR is available: {latest} "
                f"(you have v{APP_VERSION}).\n\n"
                "Open the download page?",
                "Capture OCR - update available",
                0x4 | 0x20 | 0x1000 | 0x40000)  # Yes/No, question
            if r == 6:  # IDYES
                os.startfile(page)
    except Exception:
        pass


def main():
    _cleanup_temp()
    if UPDATE_CHECK:
        import threading
        threading.Thread(target=_check_updates, daemon=True).start()
    import shutil
    if not os.path.exists(TESSERACT_EXE) and shutil.which("tesseract") is None:
        _msgbox("Tesseract OCR not found.\n\n"
                "Install Tesseract for Windows (UB Mannheim build) with the "
                "ITALIAN language data, default path:\n"
                r"C:\Program Files\Tesseract-OCR" "\n\n"
                "Without Tesseract the rectangle selection works but the "
                "OCR will fail.", icon=0x30)
    kb = keyboard.Listener(on_press=on_key_press, on_release=on_key_release)
    kb.daemon = True
    kb.start()
    mouse_listener.daemon = True
    mouse_listener.start()
    outputs = {"clipboard": "the text is copied to the clipboard "
                            "(paste it anywhere)",
               "word": "the text opens in Word",
               "both": "the text is copied to the clipboard and opens "
                       "in Word"}
    _msgbox("Capture OCR is running in the background.\n\n"
            "Usage:\n"
            "\u2022 hold ALT + RIGHT mouse button and drag = select the "
            "rectangle\n"
            "\u2022 release the right button: OCR runs and " +
            outputs.get(OUTPUT_MODE, outputs["both"]) + "\n"
            "\u2022 CTRL + ALT + Q = quit the program (with confirmation)",
            title="Capture OCR - started", icon=0x40)  # 0x40 = info
    App().run()
    os._exit(0)


if __name__ == "__main__":
    main()
