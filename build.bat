@echo off
echo =========================================
echo Building Spelling Injector...
echo =========================================

REM 1. Clean previous builds
echo [1/5] Cleaning old builds...
if exist "Spelling-Injector-Windows.zip" del /q "Spelling-Injector-Windows.zip"
if exist "Spelling-Injector" rmdir /s /q "Spelling-Injector"
if exist "main.dist" rmdir /s /q "main.dist"
if exist "main.build" rmdir /s /q "main.build"

REM 2. Compile with Nuitka (Sets exe icon, includes app_icon.ico for Tkinter windows, and copies the icons folder)
echo [2/5] Compiling with Nuitka...
py -3.12 -m nuitka --mingw64 --assume-yes-for-downloads --standalone --enable-plugin=tk-inter --windows-console-mode=disable --windows-icon-from-ico=app_icon.ico --include-data-files=app_icon.ico=app_icon.ico --include-package=pyttsx3 --include-package=comtypes --include-data-dir=icons=icons -o "Spelling Injector.exe" main.py

REM 3. Rename the output folder
echo [3/5] Renaming output folder...
rename main.dist Spelling-Injector

REM 4. Create ZIP file (using Windows built-in tar command)
echo [4/5] Zipping the release...
tar -a -c -f Spelling-Injector-Windows.zip Spelling-Injector

REM 5. Clean up temporary Nuitka build folders
echo [5/5] Cleaning up temp files...
if exist "main.build" rmdir /s /q "main.build"
if exist "Spelling-Injector" rmdir /s /q "Spelling-Injector"

echo.
echo =========================================
echo DONE! Release zip is ready: Spelling-Injector-Windows.zip
echo =========================================
pause
