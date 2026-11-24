@echo off

:: MiMoCode + ConEmu Launcher

set "MIMO_DIR=%~dp0"
set "CONEMU_EXE=C:\Program Files\ConEmu\ConEmu64.exe"

if not exist "%CONEMU_EXE%" (
    echo Error: ConEmu not found!
    echo Path: %CONEMU_EXE%
    pause
    exit /b 1
)

cd /d "%MIMO_DIR%"
start "" "%CONEMU_EXE%" -cmd mimo.exe
