@echo off
title OpsFlow Studio Launcher
echo ==================================================
echo Starting OpsFlow Studio Docker Container...
echo ==================================================

:: Set the directory to the folder containing this script
cd /d "%~dp0"

:: Launch the Docker container
docker run -p 8501:8501 opsflow-studio

:: If the app crashes, keep the window open to read the error
echo.
echo ==================================================
echo Application stopped or crashed. See error above.
echo ==================================================
pause