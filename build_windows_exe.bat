@echo off
setlocal EnableExtensions

rem ============================================================
rem UPS Bid Analyzer - Windows build script
rem
rem Run by double-clicking this file or from PowerShell/CMD:
rem     build_windows.bat
rem
rem Keep this file in the project root beside:
rem     pyproject.toml
rem     ups-venv\
rem     src\
rem ============================================================

rem Always run from the directory containing this BAT file.
cd /d "%~dp0"

set "PROJECT_ROOT=%CD%"
set "APP_NAME=Bid_Analyzer"

set "VENV_PYTHON=%PROJECT_ROOT%\ups-venv\Scripts\python.exe"
set "SRC_DIR=%PROJECT_ROOT%\src"
set "ENTRY_SCRIPT=%SRC_DIR%\bid_analyzer\__main__.py"
set "ICON_FILE=%SRC_DIR%\bid_analyzer\resources\app_icon.ico"
set "RESOURCES_DIR=%SRC_DIR%\bid_analyzer\resources"

set "DIST_DIR=%PROJECT_ROOT%\packaging\windows"
set "WORK_DIR=%PROJECT_ROOT%\build\pyinstaller"
set "SPEC_DIR=%PROJECT_ROOT%\build\spec"

echo.
echo ============================================================
echo Building %APP_NAME%
echo ============================================================
echo Project: %PROJECT_ROOT%
echo Python:  %VENV_PYTHON%
echo Output:  %DIST_DIR%\%APP_NAME%
echo.

rem ------------------------------------------------------------
rem Validate required files and folders.
rem ------------------------------------------------------------

if not exist "%VENV_PYTHON%" (
    echo ERROR: Virtual-environment Python was not found:
    echo   %VENV_PYTHON%
    echo.
    echo Update VENV_PYTHON in this BAT file if your environment
    echo has a different folder name.
    goto :failure
)

if not exist "%ENTRY_SCRIPT%" (
    echo ERROR: Entry script was not found:
    echo   %ENTRY_SCRIPT%
    goto :failure
)

if not exist "%ICON_FILE%" (
    echo ERROR: Application icon was not found:
    echo   %ICON_FILE%
    goto :failure
)

if not exist "%RESOURCES_DIR%" (
    echo ERROR: Resources directory was not found:
    echo   %RESOURCES_DIR%
    goto :failure
)

rem Confirm PyInstaller is installed in the intended environment.
"%VENV_PYTHON%" -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo ERROR: PyInstaller is not installed in ups-venv.
    echo.
    echo Install it with:
    echo   "%VENV_PYTHON%" -m pip install pyinstaller
    goto :failure
)

rem ------------------------------------------------------------
rem Remove the previous application output.
rem PyInstaller also receives --clean for its internal cache.
rem ------------------------------------------------------------

if exist "%DIST_DIR%\%APP_NAME%" (
    echo Removing previous build...
    rmdir /s /q "%DIST_DIR%\%APP_NAME%"

    if exist "%DIST_DIR%\%APP_NAME%" (
        echo ERROR: The previous build could not be removed.
        echo Close the executable or any File Explorer window using:
        echo   %DIST_DIR%\%APP_NAME%
        goto :failure
    )
)

if not exist "%DIST_DIR%" mkdir "%DIST_DIR%"
if not exist "%WORK_DIR%" mkdir "%WORK_DIR%"
if not exist "%SPEC_DIR%" mkdir "%SPEC_DIR%"

rem ------------------------------------------------------------
rem Build the application.
rem Use python -m PyInstaller so the selected virtual environment
rem is guaranteed to perform the build.
rem ------------------------------------------------------------

"%VENV_PYTHON%" -m PyInstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --icon "%ICON_FILE%" ^
    --name "%APP_NAME%" ^
    --clean ^
    --paths "%SRC_DIR%" ^
    --add-data "%RESOURCES_DIR%;bid_analyzer/resources" ^
    --distpath "%DIST_DIR%" ^
    --workpath "%WORK_DIR%" ^
    --specpath "%SPEC_DIR%" ^
    "%ENTRY_SCRIPT%"

if errorlevel 1 goto :failure

if not exist "%DIST_DIR%\%APP_NAME%\%APP_NAME%.exe" (
    echo ERROR: PyInstaller finished without creating the expected EXE:
    echo   %DIST_DIR%\%APP_NAME%\%APP_NAME%.exe
    goto :failure
)

echo.
echo ============================================================
echo BUILD SUCCESSFUL
echo ============================================================
echo Executable:
echo   %DIST_DIR%\%APP_NAME%\%APP_NAME%.exe
echo.

rem Open the completed build folder.
start "" "%DIST_DIR%\%APP_NAME%"

pause
exit /b 0


:failure
echo.
echo ============================================================
echo BUILD FAILED
echo ============================================================
echo Review the messages above for the cause.
echo.
pause
exit /b 1
