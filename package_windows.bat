@echo off
REM Package the BBGRL Slides UI into a single EXE suitable for double-clicking.
REM - Builds in an isolated virtual environment
REM - Includes Flask templates and collected assets
REM - Auto-opens browser on launch
REM - Writes logs to ui_app.log next to the EXE

setlocal
set NAME=BBGRL Slides App
set ENTRY=ui_app\app.py
set VENV_DIR=.venv-pkg

echo Creating isolated build environment: %VENV_DIR%
python -m venv %VENV_DIR%
if errorlevel 1 (
  echo Failed to create venv. Ensure Python is installed and on PATH.
  exit /b 1
)

call %VENV_DIR%\Scripts\activate
if errorlevel 1 (
  echo Failed to activate venv.
  exit /b 1
)

python -m pip install -U pip wheel
if errorlevel 1 goto :pip_fail

python -m pip install -r requirements.txt
if errorlevel 1 goto :pip_fail

python -m pip install pyinstaller
if errorlevel 1 goto :pip_fail

echo Building executable with PyInstaller...
pyinstaller --noconfirm --onefile ^
  --name "%NAME%" ^
  --add-data "ui_app\templates;templates" ^
  --add-data "bbgrl;bbgrl" ^
  --collect-all flask ^
  --collect-all jinja2 ^
  --collect-all werkzeug ^
  --collect-all itsdangerous ^
  --collect-all markupsafe ^
  --collect-all selenium ^
  --collect-all bs4 ^
  --collect-all lxml ^
  --collect-all pptx ^
  --collect-all requests ^
  "%ENTRY%"
if errorlevel 1 (
  echo Build failed.
  goto :end
)

echo.
echo Build complete. Find the EXE in the dist folder:
echo   dist\%NAME%.exe
echo Double-click it; your browser should open automatically.
echo If it does not, run from terminal to see logs:
echo   dist\"%NAME%.exe"
echo.
goto :end

:pip_fail
echo Failed to install dependencies. Try:
echo   python -m pip install -r requirements.txt
goto :end

:end
endlocal
