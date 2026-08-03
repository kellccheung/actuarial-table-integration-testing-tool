@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================================
echo  Prophet Table Change Tool - one-time Conda setup
echo ============================================================
echo.

call :find_conda
if errorlevel 1 (
  echo Conda was not found.
  echo Install Anaconda/Miniconda, or open "Anaconda Prompt" and run this Setup.bat from there.
  echo.
  pause
  exit /b 1
)

echo Using: %CONDA_EXE%
echo Creating/updating Conda environment "prophet-table" from environment.yml ...
echo.

call "%CONDA_EXE%" env update -f "%~dp0environment.yml" --prune
if errorlevel 1 (
  echo.
  echo Setup FAILED. See messages above.
  pause
  exit /b 1
)

echo.
echo Setup OK. You can now double-click Run.bat to start the tool.
echo.
pause
exit /b 0

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
