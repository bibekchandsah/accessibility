@echo off
title Camera Controller
echo Camera Device Controller
echo =======================
echo.
echo Choose an option:
echo 1. Run Camera Controller (Auto-elevate to Admin)
echo 2. Run in Limited Mode (No Admin)
echo 3. Exit
echo.
set /p choice="Enter your choice (1-3): "

if "%choice%"=="1" (
    echo Starting Camera Controller with auto-elevation...
    python advanced_camera_controller.py
) else if "%choice%"=="2" (
    echo Starting Camera Controller in Limited Mode...
    python advanced_camera_controller.py --skip-admin
) else if "%choice%"=="3" (
    exit
) else (
    echo Invalid choice. Please try again.
    pause
    goto :eof
)

pause