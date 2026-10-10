@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================================
echo  Prophet Table Change Tool - one-time Conda setup
echo ============================================================
echo.

call "%~dp0find_conda.bat"
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
