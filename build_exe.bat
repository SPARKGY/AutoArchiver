@echo off
echo Installing PyInstaller...
pip install pyinstaller

echo Building AutoArchiver executable...
REM --noconfirm: overwrite output directory
REM --onefile: bundle everything into a single .exe
REM --windowed: no console window when running
REM --icon: use the sparkgy.ico
REM --name: output name
REM --add-data: include the icon file inside the exe for runtime access
REM --hidden-import: ensure pymupdf and others are caught if missed

pyinstaller AutoArchiver.spec

echo.
echo Build Complete!
echo The executable is located in the "dist" folder:
echo .\dist\AutoArchiver.exe
echo Copying config.json to dist...
copy config.json dist\config.json
pause
