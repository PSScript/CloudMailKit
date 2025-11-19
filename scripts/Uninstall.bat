@echo off
echo ========================================
echo CloudMailReader COM Unregistration
echo ========================================
echo.

REM Check for admin rights
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script requires Administrator privileges
    echo Please run as Administrator
    pause
    exit /b 1
)

REM Get script directory
set SCRIPT_DIR=%~dp0

echo Unregistering CloudMailReader.dll...
echo.

REM Use 64-bit RegAsm for 64-bit systems
if exist "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" (
    "%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" "%SCRIPT_DIR%CloudMailReader.dll" /unregister
) else (
    "%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" "%SCRIPT_DIR%CloudMailReader.dll" /unregister
)

if %errorLevel% equ 0 (
    echo.
    echo ========================================
    echo Unregistration successful!
    echo ========================================
) else (
    echo.
    echo ========================================
    echo Unregistration failed
    echo ========================================
)

pause