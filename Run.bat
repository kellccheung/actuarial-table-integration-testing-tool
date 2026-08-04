@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================================
echo  Prophet Table Change Tool
echo ============================================================
echo.
echo Set "mode" in Control.xlsx Config sheet, then provide Control.xlsx below.
echo   generate_changelog  - Stage 1: create Change Log
echo   validate_only       - Stage 2: dry-run validation
echo   apply               - Stage 2: write New_Production_Tables
echo.

call :find_conda
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
REM strip accidental quotes from set /p
set "CONTROL=%CONTROL:"=%"

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

echo.
echo Running with Control: %CONTROL%
echo.

call "%CONDA_EXE%" run --no-capture-output -n prophet-table python "%~dp0run_launcher.py" "%CONTROL%"
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

:find_conda
where conda >nul 2>&1
if not errorlevel 1 (
  for /f "delims=" %%I in ('where conda') do (
    set "CONDA_EXE=%%I"
    exit /b 0
  )
)
if exist "%USERPROFILE%\anaconda3\Scripts\conda.exe" (
  set "CONDA_EXE=%USERPROFILE%\anaconda3\Scripts\conda.exe"
  exit /b 0
)
if exist "%USERPROFILE%\miniconda3\Scripts\conda.exe" (
  set "CONDA_EXE=%USERPROFILE%\miniconda3\Scripts\conda.exe"
  exit /b 0
)
if exist "%USERPROFILE%\Miniconda3\Scripts\conda.exe" (
  set "CONDA_EXE=%USERPROFILE%\Miniconda3\Scripts\conda.exe"
  exit /b 0
)
if exist "%LOCALAPPDATA%\anaconda3\Scripts\conda.exe" (
  set "CONDA_EXE=%LOCALAPPDATA%\anaconda3\Scripts\conda.exe"
  exit /b 0
)
if exist "%LOCALAPPDATA%\miniconda3\Scripts\conda.exe" (
  set "CONDA_EXE=%LOCALAPPDATA%\miniconda3\Scripts\conda.exe"
  exit /b 0
)
if exist "%ProgramData%\anaconda3\Scripts\conda.exe" (
  set "CONDA_EXE=%ProgramData%\anaconda3\Scripts\conda.exe"
  exit /b 0
)
if exist "%ProgramData%\miniconda3\Scripts\conda.exe" (
  set "CONDA_EXE=%ProgramData%\miniconda3\Scripts\conda.exe"
  exit /b 0
)
exit /b 1
