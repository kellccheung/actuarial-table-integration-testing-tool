@echo off
REM Thin wrapper: use project-root Run.bat (do not copy .py files here).
setlocal EnableExtensions
set "ROOT=%~dp0..\..\..\"
set "CONTROL=%~1"
if "%CONTROL%"=="" set "CONTROL=%~dp0Control.xlsx"
call "%ROOT%Run.bat" "%CONTROL%"
exit /b %ERRORLEVEL%
