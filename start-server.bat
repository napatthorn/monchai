@echo off
setlocal
cd /d %~dp0

start "Monchai Insurance Server" cmd /k npm start

timeout /t 2 >nul

start "" http://127.0.0.1:3000