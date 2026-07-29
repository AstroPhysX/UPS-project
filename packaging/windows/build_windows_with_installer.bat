@echo off
setlocal EnableExtensions

rem ============================================================
rem UPS Bid Analyzer - Windows application and installer build
rem
rem File location:
rem     packaging\windows\build_windows_with_installer.bat
rem ============================================================

set "WINDOWS_PACKAGING_DIR=%~dp0"

if "%WINDOWS_PACKAGING_DIR:~-1%"=="\" (
    set "WINDOWS_PACKAGING_DIR=%WINDOWS_PACKAGING_DIR:~0,-1%"
)

rem packaging\windows is two levels below the project root.
for %%I in ("%WINDOWS_PACKAGING_DIR%\..\..") do set "PROJECT_ROOT=%%~fI"

set "APP_NAME=Bid_Analyzer"

set "VENV_PYTHON=%PROJECT_ROOT%\ups-venv\Scripts\python.exe"
set "SRC_DIR=%PROJECT_ROOT%\src"
set "ENTRY_SCRIPT=%SRC_DIR%\bid_analyzer\__main__.py"
set "ICON_FILE=%SRC_DIR%\bid_analyzer\resources\app_icon.ico"
set "RESOURCES_DIR=%SRC_DIR%\bid_analyzer\resources"

set "DIST_DIR=%WINDOWS_PACKAGING_DIR%"
set "WORK_DIR=%PROJECT_ROOT%\build\pyinstaller"
set "SPEC_DIR=%PROJECT_ROOT%\build\spec"

set "INNO_SCRIPT=%WINDOWS_PACKAGING_DIR%\installer_windows.iss"
set "INSTALLER_DIR=%WINDOWS_PACKAGING_DIR%\installer"
set "ISCC="

cd /d "%PROJECT_ROOT%"

rem ------------------------------------------------------------
rem Locate Inno Setup.
rem ------------------------------------------------------------

if exist "%ProgramFiles%\Inno Setup 7\ISCC.exe" (
    set "ISCC=%ProgramFiles%\Inno Setup 7\ISCC.exe"
)

if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe" (
    set "ISCC=%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe"
)

if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" (
    set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
)

if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" (
    set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
)

echo.
echo ============================================================
echo Building %APP_NAME%
echo ============================================================
echo Project:
echo   %PROJECT_ROOT%
echo.
echo Application output:
echo   %DIST_DIR%\%APP_NAME%
echo.
echo Installer output:
echo   %INSTALLER_DIR%
echo.

rem ------------------------------------------------------------
rem Validate required files and tools.
rem ------------------------------------------------------------

if not exist "%VENV_PYTHON%" (
    echo ERROR: Virtual-environment Python was not found:
    echo   %VENV_PYTHON%
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

if not exist "%INNO_SCRIPT%" (
    echo ERROR: Inno Setup script was not found:
    echo   %INNO_SCRIPT%
    goto :failure
)

if not defined ISCC (
    echo ERROR: The Inno Setup compiler ISCC.exe was not found.
    echo Install Inno Setup 6 or 7, or update the ISCC paths.
    goto :failure
)

"%VENV_PYTHON%" -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo ERROR: PyInstaller is not installed in ups-venv.
    echo Install it with:
    echo   "%VENV_PYTHON%" -m pip install pyinstaller
    goto :failure
)

rem ------------------------------------------------------------
rem Remove the previous PyInstaller output.
rem ------------------------------------------------------------

if exist "%DIST_DIR%\%APP_NAME%" (
    echo Removing previous application build...
    rmdir /s /q "%DIST_DIR%\%APP_NAME%"

    if exist "%DIST_DIR%\%APP_NAME%" (
        echo ERROR: The previous build could not be removed.
        echo Close the application and any Explorer window using:
        echo   %DIST_DIR%\%APP_NAME%
        goto :failure
    )
)

if not exist "%DIST_DIR%" mkdir "%DIST_DIR%"
if not exist "%WORK_DIR%" mkdir "%WORK_DIR%"
if not exist "%SPEC_DIR%" mkdir "%SPEC_DIR%"

rem ------------------------------------------------------------
rem Build the PyInstaller --onedir application.
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
    echo ERROR: PyInstaller finished without creating:
    echo   %DIST_DIR%\%APP_NAME%\%APP_NAME%.exe
    goto :failure
)

rem ------------------------------------------------------------
rem Build the Inno Setup installer.
rem ------------------------------------------------------------

echo.
echo Building installer with:
echo   %ISCC%
echo.

pushd "%WINDOWS_PACKAGING_DIR%"
"%ISCC%" "%INNO_SCRIPT%"
set "INNO_EXIT_CODE=%ERRORLEVEL%"
popd

if not "%INNO_EXIT_CODE%"=="0" goto :failure

set "INSTALLER_EXE="
for /f "delims=" %%F in ('dir /b /a-d /o-d "%INSTALLER_DIR%\Bid_Analyzer_Setup_*.exe" 2^>nul') do (
    if not defined INSTALLER_EXE set "INSTALLER_EXE=%INSTALLER_DIR%\%%F"
)

if not defined INSTALLER_EXE (
    echo ERROR: Inno Setup finished without creating:
    echo   %INSTALLER_DIR%\Bid_Analyzer_Setup_*.exe
    goto :failure
)

echo.
echo ============================================================
echo BUILD SUCCESSFUL
echo ============================================================
echo Application:
echo   %DIST_DIR%\%APP_NAME%\%APP_NAME%.exe
echo.
echo Installer:
echo   %INSTALLER_EXE%
echo.

start "" "%INSTALLER_DIR%"

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
