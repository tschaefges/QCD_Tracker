@echo off
echo Building lookup_charity.exe with PyInstaller...
echo.
pyinstaller --onefile --console --name lookup_charity lookup_charity.py
echo.
if exist dist\lookup_charity.exe (
    echo Build complete. EXE is at: dist\lookup_charity.exe
    echo.
    echo Copy these two files to each user's QCD folder:
    echo   dist\lookup_charity.exe
    echo   config.txt
) else (
    echo Build may have failed. Check output above for errors.
)
pause
