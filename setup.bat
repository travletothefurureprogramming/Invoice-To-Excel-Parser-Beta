@echo off
setlocal
title Invoice to Excel Parser - Launcher

:: ============================================================
:: Invoice to Excel Parser
:: Setup & Launcher
:: ============================================================

cd /d "%~dp0"

echo.
echo ============================================================
echo          INVOICE TO EXCEL PARSER
echo                 BETA VERSION
echo ============================================================
echo.
echo [INFO] Project folder:
echo        %CD%
echo.

:: ------------------------------------------------------------
:: Check Python
:: ------------------------------------------------------------

echo [1/5] Checking Python...

python --version >nul 2>&1

if errorlevel 1 (
    echo.
    echo [ERROR] Python was not found.
    echo.
    echo Please install Python 3.10 or newer and make sure
    echo "Add Python to PATH" is enabled during installation.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%A in ('python --version') do set PYTHON_VERSION=%%A

echo [OK] %PYTHON_VERSION%
echo.

:: ------------------------------------------------------------
:: Create virtual environment
:: ------------------------------------------------------------

echo [2/5] Checking virtual environment...

if not exist ".venv\Scripts\python.exe" (
    echo [INFO] Creating virtual environment...
    python -m venv .venv

    if errorlevel 1 (
        echo.
        echo [ERROR] Failed to create virtual environment.
        echo.
        pause
        exit /b 1
    )

    echo [OK] Virtual environment created.
) else (
    echo [OK] Virtual environment already exists.
)

echo.

:: ------------------------------------------------------------
:: Activate virtual environment
:: ------------------------------------------------------------

echo [3/5] Activating virtual environment...

call ".venv\Scripts\activate.bat"

if errorlevel 1 (
    echo.
    echo [ERROR] Failed to activate virtual environment.
    echo.
    pause
    exit /b 1
)

echo [OK] Virtual environment activated.
echo.

:: ------------------------------------------------------------
:: Install requirements
:: ------------------------------------------------------------

echo [4/5] Installing dependencies...
echo.

python -m pip install --upgrade pip

if errorlevel 1 (
    echo.
    echo [WARNING] Could not upgrade pip.
    echo [INFO] Continuing anyway...
    echo.
)

if not exist "requirements.txt" (
    echo.
    echo [ERROR] requirements.txt was not found.
    echo.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ============================================================
    echo [ERROR] Failed to install one or more dependencies.
    echo ============================================================
    echo.
    echo Check requirements.txt and your internet connection.
    echo.
    pause
    exit /b 1
)

echo.
echo [OK] Dependencies installed.
echo.

:: ------------------------------------------------------------
:: Run application
:: ------------------------------------------------------------

echo [5/5] Starting Invoice to Excel Parser...
echo.
echo ============================================================
echo                 STARTING APPLICATION
echo ============================================================
echo.

if not exist "main.py" (
    echo.
    echo [ERROR] main.py was not found.
    echo.
    pause
    exit /b 1
)

python main.py

if errorlevel 1 (
    echo.
    echo ============================================================
    echo [ERROR] The application exited with an error.
    echo ============================================================
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo              APPLICATION CLOSED
echo ============================================================
echo.

pause
endlocal