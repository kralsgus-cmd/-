@echo off
cd /d "%~dp0"
python --version >nul 2>&1
if errorlevel 1 (
    echo Python not found. Please install from python.org
    pause
    exit
)
python teramath.py
if errorlevel 1 pause
