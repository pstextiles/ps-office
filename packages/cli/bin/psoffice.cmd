@echo off
rem psoffice launcher for the packaged Windows app: <install>\resources\cli\psoffice.cmd
setlocal
set ELECTRON_RUN_AS_NODE=1
"%~dp0..\..\PSOffice.exe" "%~dp0genoffice.cjs" %*
endlocal
