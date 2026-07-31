@echo off
title OpsFlow Studio Shutdown
echo ==================================================
echo Shutting down OpsFlow Studio...
echo ==================================================

echo.
echo Stopping and cleaning up ALL OpsFlow containers...
:: Stop and remove the explicitly named container
docker stop opsflow-container 2>nul
docker rm opsflow-container 2>nul

:: Find and force-remove any "ghost" containers built from the opsflow-studio image
FOR /f "tokens=*" %%i IN ('docker ps -a -q -f ancestor^=opsflow-studio') DO docker rm -f %%i 2>nul

echo.
echo Closing the Launcher window...
:: This hunts down the RunOpsFlow window by its specific title and closes it automatically
taskkill /F /FI "WINDOWTITLE eq OpsFlow Studio Launcher*" 2>nul

echo.
echo ==================================================
echo ✅ OpsFlow Studio has been successfully closed and cleaned.
echo ==================================================
:: This will wait 3 seconds so you can read the message, then close this window automatically
timeout /t 3 >nul