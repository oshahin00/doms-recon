@echo off
title OpsFlow Studio Launcher
echo ==================================================
echo Preparing OpsFlow Studio Environment...
echo ==================================================

:: Set the directory to the folder containing this script
cd /d "%~dp0"

:: Create the output folder on your desktop if it doesn't exist yet
if not exist output mkdir output

echo.
echo 1/3: Cleaning up ALL old OpsFlow containers...
:: Stop and remove the explicitly named container
docker stop opsflow-container 2>nul
docker rm opsflow-container 2>nul
:: Find and force-remove any "ghost" containers built from the opsflow-studio image
FOR /f "tokens=*" %%i IN ('docker ps -a -q -f ancestor^=opsflow-studio') DO docker rm -f %%i 2>nul

echo.
echo 2/3: Building the latest Docker image...
docker build -t opsflow-studio .

echo.
echo ==================================================
echo 3/3: Starting OpsFlow Studio...
echo ==================================================
:: Launch Docker with a specific name, timezone set to local, and the volume mount
docker run --name opsflow-container -e TZ=Asia/Amman -p 8501:8501 -v "%cd%\output:/app/output" opsflow-studio

:: If the app crashes, keep the window open to read the error
echo.
echo ==================================================
echo Application stopped or crashed. See error above.
echo ==================================================
pause