@echo off
setlocal
cd /d "%~dp0"
set "BOM_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "BOM_PYW=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"
if not exist "%BOM_PY%" (
  echo Codex Python runtime was not found. See README.md in this folder.
  pause
  exit /b 1
)
if not "%~1"=="" (
  "%BOM_PY%" "%~dp0app.py" %*
  exit /b
)
start "" "%BOM_PYW%" "%~dp0app.py"
exit /b 0
