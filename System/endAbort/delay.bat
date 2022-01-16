@echo off
taskkill /f /im wscript.exe
timeout 2 /nobreak /t >nul
start done.vbs
