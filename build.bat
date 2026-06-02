@echo off
REM ============================================================
REM  Illinois Tax Scraper - Build Setup EXE
REM  Run this once from the project folder to produce:
REM    dist\TaxScraperSetup.exe
REM ============================================================

echo.
echo ==========================================
echo  Building Illinois Tax Scraper Setup EXE
echo ==========================================
echo.

REM Make sure we are in the right directory
cd /d "%~dp0"

REM Install PyInstaller if missing
python -m pip install pyinstaller --quiet

REM Clean previous build
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

REM Build the EXE
REM   --onefile          : Single .exe file
REM   --noconsole        : No black terminal window (GUI only)
REM   --name             : Output EXE name
REM   --add-data         : Bundle tax_scraper.py and the Excel file
REM   --hidden-import    : Ensure tkinter submodules are included

python -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --name "TaxScraperSetup" ^
  --add-data "tax_scraper.py;." ^
  --add-data "Tax il.xlsm;." ^
  --hidden-import "tkinter" ^
  --hidden-import "tkinter.ttk" ^
  --hidden-import "tkinter.filedialog" ^
  --hidden-import "tkinter.messagebox" ^
  setup_installer.py

echo.
if exist "dist\TaxScraperSetup.exe" (
    echo ==========================================
    echo  SUCCESS!
    echo  EXE created: dist\TaxScraperSetup.exe
    echo.
    echo  Distribute to client:
    echo    1. dist\TaxScraperSetup.exe
    echo    (Tax il.xlsm is bundled inside the EXE)
    echo ==========================================
) else (
    echo ==========================================
    echo  BUILD FAILED - check output above
    echo ==========================================
)

pause
