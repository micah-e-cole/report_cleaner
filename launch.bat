@echo off
REM ================================
REM Classroom Utilization Cleaner
REM ================================

REM Change directory to the folder where this .bat file lives
cd /d "%~dp0"

echo.
echo ===============================================
echo   Checking for Python installation...
echo ===============================================

REM -----------------------------------------------
REM 1. Detect Python ("py" preferred, fallback "python")
REM -----------------------------------------------
where py >nul 2>&1
if %errorlevel%==0 (
    set "PY_CMD=py"
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        set "PY_CMD=python"
    ) else (
        echo.
        echo ERROR: Python is not installed or not on PATH.
        echo Please install Python and try again.
        echo.
        pause
        exit /b 1
    )
)

echo Found Python command: %PY_CMD%
echo.

REM -----------------------------------------------
REM 2. Create virtual environment if needed
REM    (and remember whether we just created it)
REM -----------------------------------------------
set "JUST_CREATED_VENV=0"

if not exist ".venv" (
    echo Creating virtual environment in .venv ...
    %PY_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo.
        echo ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
    set "JUST_CREATED_VENV=1"
) else (
    echo Virtual environment already exists. Skipping creation.
)

echo.

REM -----------------------------------------------
REM 3. Activate the virtual environment
REM -----------------------------------------------
call ".venv\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Could not activate virtual environment.
    pause
    exit /b 1
)

echo Virtual environment activated.
echo.

REM -----------------------------------------------
REM 4. Install dependencies (only on first creation)
REM -----------------------------------------------
if "%JUST_CREATED_VENV%"=="1" (
    if exist requirements.txt (
        echo First-time setup: installing dependencies from requirements.txt ...
        %PY_CMD% -m pip install --upgrade pip
        %PY_CMD% -m pip install -r requirements.txt
        if %errorlevel% neq 0 (
            echo.
            echo ERROR: Failed to install Python dependencies.
            pause
            exit /b 1
        )
    ) else (
        echo WARNING: requirements.txt not found. Skipping dependency install.
    )
) else (
    echo Using existing virtual environment. Skipping dependency installation.
)

echo.
echo Environment ready.
echo.

REM -----------------------------------------------
REM 5. Run the GUI
REM -----------------------------------------------
echo Launching the Classroom Utilization Cleaner...
%PY_CMD% "%~dp0run_gui.py"

REM If Python returned a non-zero exit code, keep window open for debugging
if %errorlevel% neq 0 (
    echo.
    echo The program encountered an error. See details above.
    echo.
    pause
) else (
    REM Normal exit, close window
    exit /b
)