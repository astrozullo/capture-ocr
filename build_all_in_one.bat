@echo off
setlocal enabledelayedexpansion
echo =====================================================
echo  Building CaptureOCR.exe ALL-IN-ONE
echo  (Python and Tesseract inside a single executable)
echo =====================================================
echo.
echo 0) Looking for Tesseract on this PC...
set "TESS="
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" set "TESS=C:\Program Files\Tesseract-OCR"
if not defined TESS if exist "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" set "TESS=C:\Program Files (x86)\Tesseract-OCR"
if not defined TESS if exist "%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe" set "TESS=%LOCALAPPDATA%\Programs\Tesseract-OCR"
if not defined TESS if exist "%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe" set "TESS=%LOCALAPPDATA%\Tesseract-OCR"
if not defined TESS (
    for /f "delims=" %%i in ('where tesseract 2^>nul') do if not defined TESS set "TESS=%%~dpi"
)
if not defined TESS (
    echo ERROR: Tesseract not found in any known location.
    echo Find tesseract.exe yourself ^(Start menu, type tesseract,
    echo right click, "Open file location"^) and type below the
    echo folder that contains it, then press Enter:
    set /p "TESS=Tesseract folder: "
)
if "%TESS:~-1%"=="\" set "TESS=%TESS:~0,-1%"
if not exist "%TESS%\tesseract.exe" (
    echo ERROR: no tesseract.exe in "%TESS%". Try again.
    pause
    exit /b 1
)
echo    Found: %TESS%
echo.
echo 1) Installing/updating the required Python packages...
py -m pip install --upgrade pip >nul
py -m pip install pynput mss pytesseract python-docx pillow pywin32 pyinstaller
echo.
echo 2) Copying Tesseract into the working folder...
robocopy "%TESS%" "Tesseract-OCR" /E /NFL /NDL /NJH /NJS >nul
if errorlevel 8 (
    echo ERROR while copying Tesseract.
    pause
    exit /b 1
)
echo 3) Building (a few minutes, the exe will also contain Tesseract)...
py -m PyInstaller --onefile --noconsole --name CaptureOCR --collect-data docx --hidden-import win32timezone --add-data "Tesseract-OCR;Tesseract-OCR" capture_ocr.pyw
echo.
if exist dist\CaptureOCR.exe (
    echo DONE: dist\CaptureOCR.exe is SELF-CONTAINED.
    echo On any PC, just copy that single file:
    echo no Python, no Tesseract, nothing to download.
) else (
    echo ERROR: executable not created, read the messages above.
)
pause
