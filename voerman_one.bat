@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Voerman OneBAT
pushd "%~dp0"

echo [INFO] Project: %CD%
if not exist ".env" ( if exist ".env.example" copy /Y ".env.example" ".env" >nul )
if not exist ".venv" (
  echo [INFO] Creating .venv ...
  where py >nul 2>nul && (py -3 -m venv .venv) || (python -m venv .venv)
)
echo [INFO] Installing/Checking packages ...
call ".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
call ".venv\Scripts\python.exe" -m pip install -r requirements.txt

if /I "%~1"=="demo" goto :demo
if /I "%~1"=="dev" goto :dev
if /I "%~1"=="debug" goto :debug

set PYTHONPATH=%CD%
if not exist "out" mkdir "out"
del /f /q "out\port.txt" >nul 2>nul

echo [INFO] Starting server (inline). Dit venster open laten.
call ".venv\Scripts\python.exe" "%CD%\run_server.py"
echo [INFO] Server gestopt of gecrasht. (Zie melding hierboven.) Druk op een toets om te sluiten.
pause >nul
goto :end

:dev
set PYTHONPATH=%CD%
echo [DEV] Uvicorn met reload
call ".venv\Scripts\python.exe" -m uvicorn --app-dir "%CD%" app:app --host 127.0.0.1 --port 8000 --reload --reload-dir "%CD%"
goto :end

:debug
set PYTHONPATH=%CD%
if not exist "out" mkdir "out"
echo [DEBUG] run_server.py zonder browser-open
call ".venv\Scripts\python.exe" "%CD%\run_server.py"
goto :end

:demo
call ".venv\Scripts\python.exe" "tests\run_e2e_demo.py"
if not exist "out" mkdir "out"
start "" "out"
goto :end

:end
popd
endlocal
