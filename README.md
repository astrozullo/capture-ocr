# Capture OCR → Word

A background utility for Windows: select a rectangle anywhere on your screen with **ALT + right mouse button**, the program runs OCR (Italian + English by default) on its content and delivers the recognized text, ready to use. Line breaks are preserved.

By default the text is delivered **both ways at once**: copied to the clipboard (plain text, paste it anywhere) and opened in a new, unsaved Microsoft Word document, where words the OCR is unsure about are shown in **red**. No file is written to disk: you decide whether to save. The `OUTPUT_MODE` setting switches to clipboard-only or Word-only.

## Usage

| Action | Command |
|---|---|
| Select the region | hold **ALT + RIGHT mouse button** and drag |
| Run the OCR | release the right button |
| Quit the program | **CTRL + ALT + Q** (asks for confirmation) |

Behavior in detail:

- On startup, an info box confirms the program is running in the background and recaps the commands.
- During selection, the screen is covered by a dark translucent veil and the rectangle is drawn in red. The right click is suppressed while capturing, so the context menu of the underlying application does not open: this is intentional.
- When the right button is released, the OCR runs (1–3 seconds) and the text is delivered according to `OUTPUT_MODE` (default `"both"`): plain text on the **clipboard**, and/or a **new, unsaved** Word document, as if you had typed it yourself. Save it from Word only if you need it; close without saving and nothing is left behind.
- If Word cannot be reached via COM automation, the program falls back to creating a `.docx` file in a temporary folder (`%TEMP%\CaptureOCR`, cleaned automatically of files older than 2 days) and opens it from there. In `"both"` mode, if Word is missing entirely, the clipboard delivery still succeeds.
- The clipboard copy is plain text: the red marking of doubtful words exists only in the Word output.
- **Red words** mean "check this", not "certainly wrong": they are words whose OCR confidence is below the threshold (default 60 on a 0–100 scale). Confident words stay black.
- A confirmation sound plays on success; a warning box appears if no text is found in the rectangle; errors show an on-screen alert and are appended to `capture_ocr_errors.log` in your user folder (`C:\Users\<you>\`).
- If Tesseract is not found at startup, a warning box says so and explains what to install.

## Requirements

**To run the all-in-one executable** (the normal way to use it):

- Windows 10/11
- Microsoft Word, only for the Word output (with the default `OUTPUT_MODE = "both"`, the clipboard delivery works even without Word)

Nothing else: the executable embeds Python, all libraries and Tesseract with its language data.

**To build the executable, or to run the script directly**:

- Windows 10/11
- [Python 3](https://www.python.org) (check "Add Python to PATH" during installation)
- [Tesseract OCR for Windows](https://github.com/UB-Mannheim/tesseract) (UB Mannheim build), installed with the *Italian* additional language data (English is always included)
- Python packages: `pynput mss pytesseract python-docx pillow pywin32` (the build script installs them for you)

## Building the all-in-one executable (one time only)

1. Put `capture_ocr.pyw` and `build_all_in_one.bat` in the same folder, on a PC that meets the build requirements above.
2. Double-click `build_all_in_one.bat`. It looks for Tesseract on its own (Program Files, Program Files (x86), per-user installs in AppData, then the system PATH; if it still cannot find it, it asks you for the folder), installs the required packages, copies Tesseract next to the script and builds with PyInstaller.
3. When it finishes, you get **`dist\CaptureOCR.exe`**.

That single file is all you need to distribute. The program looks for Tesseract in this order: embedded inside the executable → a `Tesseract-OCR` folder next to the executable/script (portable use) → `C:\Program Files\Tesseract-OCR` → the per-user installation in `AppData\Local\Programs`.

### Notes on the executable

- On first launch on a new PC, Windows SmartScreen or the antivirus may ask for confirmation because the executable is not digitally signed: "More info" → "Run anyway". This is standard for unsigned PyInstaller executables.
- The executable is large (on the order of 100 MB, since it embeds Python and Tesseract) and takes a few seconds to start, because it self-extracts to a temporary folder. Once running, it stays in the background with negligible resource usage.
- Auto-start with Windows: put a shortcut to the executable in the startup folder (Win+R → `shell:startup` → Enter).
- Do not commit the build output: `build/`, `dist/`, the local `Tesseract-OCR/` copy and the `.spec` file are ignored via `.gitignore`. To distribute the executable from GitHub, attach it to a *Release*.

## Running from source

Double-click `capture_ocr.pyw` (the `.pyw` extension launches it silently, with no console window). Requires the build requirements above. To install the packages manually:

```
py -m pip install pynput mss pytesseract python-docx pillow pywin32
```

## Configuration

All settings live in the `CONFIGURATION` section at the top of `capture_ocr.pyw`. Edit them **before** building the executable.

| Setting | Default | Meaning |
|---|---|---|
| `APP_VERSION` | `"1.0.0"` | this copy's version; bump it before building each release |
| `UPDATE_CHECK` | `True` | check GitHub Releases for a newer version at startup (see Updates); `False` = no network access at all |
| `GITHUB_REPO` | `"astrozullo/capture-ocr"` | the `owner/repo` whose Releases are checked |
| `TESSERACT_EXE` | `C:\Program Files\Tesseract-OCR\tesseract.exe` | standard Tesseract path (the lookup order above can override it) |
| `OCR_LANGUAGES` | `"ita+eng"` | Tesseract languages; e.g. `"eng"` for English only |
| `OUTPUT_MODE` | `"both"` | where the text goes: `"clipboard"` (plain text on the clipboard), `"word"` (Word document with red marking), `"both"` |
| `OUTPUT_FOLDER` | the user's real Desktop | destination of `.docx` files when `SAVE_TO_DESKTOP = True` (OneDrive Desktop redirection is handled) |
| `OPEN_FILE_WHEN_DONE` | `True` | (file method) open the `.docx` as soon as it is created |
| `OPEN_WITHOUT_FILE` | `True` | (Word output) deliver the text as a new, unsaved Word document, zero files on disk; `False` switches to the `.docx` file method |
| `SAVE_TO_DESKTOP` | `False` | (file method) `True` leaves the `.docx` files on the Desktop instead of the temporary folder |
| `TEMP_CLEANUP_DAYS` | `2` | on startup, delete temporary files older than this many days |
| `DOUBT_THRESHOLD` | `60` | below this confidence a word is shown in red; raise to 70 if too many correct words are flagged, lower to 50 if errors slip through |
| `SHOW_STATUS_WIDGET` | `False` | `True` shows a small draggable translucent status window ("ready" / "done HH:MM" / "no text" / "ERROR") with a ✕ to quit; it hides itself during capture so it never appears in the screenshot |
| `MIN_SIZE_PX` | `8` | rectangles smaller than this (width or height) are ignored |
| `OVERLAY_ALPHA` | `0.30` | opacity of the dark veil during selection |

## How the OCR works

Each capture is processed in 3 variants (3× upscaled grayscale with autocontrast, binarized with Otsu thresholding, inverted binarized for light text on dark backgrounds), each with 2 Tesseract segmentation modes (psm 6 = uniform block, psm 11 = sparse text). The result with the highest mean word confidence wins, and lines are rebuilt from Tesseract's word data so line breaks match the screen.

Accuracy is good on normal on-screen text (documents, web pages, PDFs, business software). On logos with fancy typefaces and on handwriting, Tesseract can fail; in those cases the words are flagged in red.

## Updates

The program carries its version in the `APP_VERSION` constant. On startup (if `UPDATE_CHECK = True`, the default), a background thread sends **one** HTTPS request to `api.github.com` asking for the latest release of `GITHUB_REPO`. If that release is newer than `APP_VERSION`, a box offers to open its download page; otherwise nothing is shown. The check is silent on any failure (offline, no releases published yet, API rate limit) and never blocks or delays the capture features.

Publishing an update (maintainer workflow) is **fully automated** by the GitHub Actions workflow in `.github/workflows/release.yml`:

1. Edit the code and bump `APP_VERSION` in `capture_ocr.pyw` (e.g. `"1.1.0"`).
2. Push the change to `main` (or edit the file directly on GitHub and commit).

That is all. The workflow runs on a GitHub-hosted Windows machine: it reads `APP_VERSION` from the script, installs Tesseract (with the Italian language data) and the Python packages, builds the all-in-one `CaptureOCR.exe` with PyInstaller, and publishes a **Release** tagged `v<APP_VERSION>` with the executable attached and auto-generated notes. If a release with that tag already exists (a push that did not bump the version), the workflow ends without doing anything, so ordinary edits never publish by accident. The workflow can also be started by hand from the Actions tab (`workflow_dispatch`).

`build_all_in_one.bat` remains as the local alternative for building on your own PC without GitHub.

Every copy of the previous version will offer the new one at its next startup. Users who prefer zero network access set `UPDATE_CHECK = False` (the program then never touches the network, but must check for updates manually on the Releases page).

## Privacy and security

- **The only network access** is the optional update check at startup described above: a single HTTPS request to `api.github.com` asking which release is the latest. No data about you, your screen or your captures is sent, ever. With `UPDATE_CHECK = False` the program makes no network access at all. This is verifiable in the source, a single readable file.
- It uses global mouse and keyboard hooks **only** to detect the two shortcuts (ALT + right click and CTRL + ALT + Q); no keystrokes are recorded or stored. Note that some antivirus tools may flag this hook mechanism, since it is the same one keyloggers use: legitimate behavior, suspicious-looking signature.
- The only disk writes are the error log and, in the file method, the `.docx` files.
- Capture happens only when you trigger it; the program takes no screenshots on its own.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| "Tesseract OCR not found" warning at startup | Install Tesseract (UB Mannheim build) with the Italian language data, or place a `Tesseract-OCR` folder next to the executable |
| SmartScreen / antivirus blocks the exe | Unsigned PyInstaller executable: "More info" → "Run anyway" (see security notes above) |
| Word does not open, a `.docx` opens instead | COM automation to Word failed; the fallback file method kicked in. Details in `capture_ocr_errors.log` |
| Confirmation beep but nothing visible on screen | `OUTPUT_MODE = "clipboard"`: the text is on the clipboard, paste it with CTRL+V |
| Poor recognition on small text | Zoom the source on screen before capturing; the doubtful words are flagged in red anyway |
| Error alert during OCR | Full traceback in `capture_ocr_errors.log` in your user folder |

## Files in this repository

| File | Purpose |
|---|---|
| `capture_ocr.pyw` | the whole program, a single Python script |
| `.github/workflows/release.yml` | GitHub Actions workflow: builds the exe and publishes the Release automatically on every version bump |
| `build_all_in_one.bat` | one-click LOCAL build of the self-contained `CaptureOCR.exe` (alternative to the workflow) |
| `README.md` | this document |
| `LICENSE` | MIT license |
| `.gitignore` | keeps build output and logs out of the repository |

## License

MIT, see [LICENSE](LICENSE).
