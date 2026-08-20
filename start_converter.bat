@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHON_EXE=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "BUNDLED_BIN=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\bin"
set "POPPLER_BIN=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin"

if not exist "%PYTHON_EXE%" (
  set "PYTHON_EXE=python"
)

if exist "%BUNDLED_BIN%" (
  set "PATH=%BUNDLED_BIN%;%PATH%"
)

if exist "%POPPLER_BIN%" (
  set "PATH=%POPPLER_BIN%;%PATH%"
)

"%PYTHON_EXE%" "%ROOT%server\converter_server.py"
pause
