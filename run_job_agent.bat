
@echo off
setlocal

REM Always run from the project folder
cd /d "%~dp0"

REM Find the project's virtual environment
set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    set "PYTHON_EXE=%~dp0venv\Scripts\python.exe"
)

REM Stop if the virtual environment was not found
if not exist "%PYTHON_EXE%" (
    echo ERROR: Virtual environment Python was not found.
    echo Expected .venv or venv inside the project folder.
    pause
    exit /b 1
)

REM Create output folder if needed
if not exist "output" mkdir "output"

REM Record the run and capture program output
echo. >> "output\scheduler.log"
echo ======================================== >> "output\scheduler.log"
echo Started: %date% %time% >> "output\scheduler.log"

"%PYTHON_EXE%" -u "%~dp0main.py" >> "output\scheduler.log" 2>&1

set "EXIT_CODE=%ERRORLEVEL%"
echo Finished: %date% %time% ^| Exit code: %EXIT_CODE% >> "output\scheduler.log"

exit /b %EXIT_CODE%
