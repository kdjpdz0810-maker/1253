@echo off
REM ===== Build PictureConverter.exe (run on Windows, Python 3.9+ required) =====
cd /d "%~dp0"
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
if errorlevel 1 goto :err

py -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name PictureConverter ^
  --collect-all tkinterdnd2 ^
  --hidden-import win32timezone ^
  --hidden-import pythoncom ^
  --hidden-import win32com.client ^
  --exclude-module matplotlib --exclude-module pandas ^
  image_converter.py
if errorlevel 1 goto :err

echo.
echo Done: dist\PictureConverter.exe
echo (HWPX access permission: put FilePathCheckerModule*.dll next to the exe)
pause
exit /b 0

:err
echo Build failed.
pause
exit /b 1
