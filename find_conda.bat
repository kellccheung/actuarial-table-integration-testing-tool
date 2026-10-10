@echo off
REM Sets CONDA_EXE for the caller. No setlocal, so the variable stays visible.
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
