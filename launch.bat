@echo off
setlocal

REM Check if the 'env' folder exists
if not exist "env" (
    echo Creating virtual environment...
    python -m venv env
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to create virtual environment.
        pause
        exit /b %ERRORLEVEL%
    )
)

REM Activate the virtual environment
call env\Scripts\activate
if %ERRORLEVEL% NEQ 0 (
    echo Failed to activate virtual environment.
    pause
    exit /b %ERRORLEVEL%
)

REM Install requirements
if exist "requirements.txt" (
    echo Installing dependencies from Requirements.txt...
    pip install -r Requirements.txt
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to install requirements.
        pause
        exit /b %ERRORLEVEL%
    )
) else (
    echo Requirements.txt not found. Skipping installation.
)

REM Display menu
:menu
cls
echo Email Extractor Launcher
echo =======================
echo 1. Simple Extractor
echo 2. Full Extractor
echo 3. Full Extractor GUI
echo 4. Advanced Extractor
echo 5. Exit
echo =======================
set /p choice="Enter your choice (1-5): "

REM Process choice
if "%choice%"=="1" (
    echo Launching Simple Extractor...
    python Simple_extractor.py
    pause
    goto menu
) else if "%choice%"=="2" (
    echo Launching Full Extractor...
    python Full_extractor.py
    pause
    goto menu
) else if "%choice%"=="3" (
    echo Launching Full Extractor GUI...
    python Full_extractor_GUI.py
    pause
    goto menu
) else if "%choice%"=="4" (
    echo Launching Advanced Extractor...
    python Advanced_Email_Extractor.py
    pause
    goto menu
) else if "%choice%"=="5" (
    echo Exiting...
    goto :end
) else (
    echo Invalid choice. Please enter a number between 1 and 5.
    pause
    goto menu
)

:end
deactivate
echo Goodbye!
pause
exit /b 0