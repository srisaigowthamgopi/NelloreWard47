@echo off
title Ward 47 Municipal Citizen Portal
cd /d "%~dp0"
echo ========================================================
echo   47th Municipal Ward Nellore Rural - Citizen Portal
echo ========================================================
echo.

where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed or not in your PATH!
    echo.
    echo To run the server and save complaints into Excel:
    echo 1. Download and install Node.js from: https://nodejs.org/
    echo 2. Restart your computer or terminal after installing.
    echo 3. Double-click this start-server.bat file again.
    echo.
    echo Press any key to open https://nodejs.org in your browser...
    pause >nul
    start https://nodejs.org/
    exit /b
)

echo [1/3] Node.js detected:
node -v
echo.

if not exist node_modules (
    echo [2/3] Installing dependencies (express, exceljs, multer)...
    call npm install
    echo.
) else (
    echo [2/3] Dependencies are ready.
    echo.
)

echo [3/3] Starting Ward 47 Citizen Portal Server...
echo.
echo Citizen Portal URL:  http://localhost:3000
echo Admin Dashboard URL: http://localhost:3000/admin.html
echo.
echo Opening Admin Portal in browser...
start http://localhost:3000/admin.html

node server.js
pause
