@echo off
REM ================================
REM Classroom Utilization Cleaner
REM ================================

REM Change directory to the folder where this .bat file lives
cd /d "%~dp0"

REM OPTIONAL: Activate virtual environment if you use one
REM Uncomment the next line if you have .venv
REM call "%~dp0.venv\Scripts\activate.bat"

REM Run the GUI
REM OR python -- depending on installation
py "%~dp0run_gui.py"

REM Keep the window open so you can see errors if something goes wrong
echo.
echo Press any key to close this window...
pause >nul
