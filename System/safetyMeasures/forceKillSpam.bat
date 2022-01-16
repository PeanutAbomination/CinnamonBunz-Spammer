@echo off
cls
color 09
echo.
echo Aborting spam.vbs, this might take some time......
taskkill /f /im wscript.exe
timeout /t 2 /nobreak > NUL
echo Spam Aborted!
start finish.vbs /wait
