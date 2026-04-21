@echo off
set "ROOT=%~dp0..\"

if exist "%ROOT%.venv\Scripts\python.exe" (
    set "PYTHON=%ROOT%.venv\Scripts\python.exe"
) else if exist "%ROOT%venv\Scripts\python.exe" (
    set "PYTHON=%ROOT%venv\Scripts\python.exe"
) else if exist "%ROOT%env\Scripts\python.exe" (
    set "PYTHON=%ROOT%env\Scripts\python.exe"
) else (
    where python >nul 2>&1
    if errorlevel 1 (
        echo Python is not installed or not found in PATH.
        exit /b 1
    )
    set "PYTHON=python"
)

where quarto >nul 2>&1
if errorlevel 1 (
    echo Quarto is not installed or not found in PATH.
    exit /b 1
)

"%PYTHON%" "%ROOT%script\main.py" %*
