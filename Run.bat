@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================================
echo  Prophet Table Change Tool
echo ============================================================
echo.
echo Set "mode" in Control.xlsx Config sheet, then provide Control.xlsx below.
echo   generate_changelog  - Stage 1: create Change Log
echo   scan                - Stage 1: check table headers only
echo   sync_requests       - add new ChangeRequests folders to Control.xlsx
echo   validate_only       - Stage 2: dry-run validation
echo   apply               - Stage 2: write New_Production_Tables
echo.

call "%~dp0find_conda.bat"
if errorlevel 1 (
  echo Conda was not found.
  echo Install Anaconda/Miniconda, or open "Anaconda Prompt" and run this from there.
  echo.
  pause
  exit /b 1
)

call "%CONDA_EXE%" run -n prophet-table python -c "import openpyxl, polars, xlsxwriter" >nul 2>&1
if errorlevel 1 (
  echo Conda environment "prophet-table" is missing or incomplete.
  echo Please double-click Setup.bat first, then try again.
  echo.
  pause
  exit /b 1
)

set "CONTROL=%~1"
if "%CONTROL%"=="" (
  set /p CONTROL=Path to Control.xlsx: 
)
REM strip accidental quotes from set /p (skip when blank: empty %VAR:"=% breaks cmd parsing)
if defined CONTROL set "CONTROL=%CONTROL:"=%"

if "%CONTROL%"=="" (
  echo No Control.xlsx path provided.
  echo Tip: you can also drag Control.xlsx onto this Run.bat file.
  echo.
  pause
  exit /b 1
)

if not exist "%CONTROL%" (
  echo File not found: %CONTROL%
  echo.
  pause
  exit /b 1
)

set "CHANGELOG=%~2"
if "%CHANGELOG%"=="" (
  echo.
  echo Change Log is only used for validate_only / apply; leave blank for newest ^(or for Stage 1^).
  set /p CHANGELOG=Change Log path or filename ^(blank = newest^): 
)
REM strip accidental quotes from set /p (skip when blank: empty %VAR:"=% breaks cmd parsing)
if defined CHANGELOG set "CHANGELOG=%CHANGELOG:"=%"

echo.
echo Running with Control: %CONTROL%
if not "%CHANGELOG%"=="" echo Change Log: %CHANGELOG%
echo.

if "%CHANGELOG%"=="" (
  call "%CONDA_EXE%" run --no-capture-output -n prophet-table python -m prophet_table_tool "%CONTROL%"
) else (
  call "%CONDA_EXE%" run --no-capture-output -n prophet-table python -m prophet_table_tool "%CONTROL%" --change-log "%CHANGELOG%"
)
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" (
  echo Finished with errors ^(exit code %RC%^).
) else (
  echo Finished successfully.
)
echo.
pause
exit /b %RC%
