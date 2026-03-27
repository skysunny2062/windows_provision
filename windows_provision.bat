@echo off
setlocal enabledelayedexpansion
COLOR 0B
net session >nul 2>&1 || goto GetAdmin
::====================================================================================
for /f "delims=" %%P in ('where python 2^>nul') do (
    if not defined PYTHON (
        echo %%P | findstr /i "WindowsApps" >nul || set "PYTHON=%%P"
    )
)
if not defined PYTHON call :PythonCheck
"!PYTHON!" %~dp0core\main.py
exit /b

::====================================================================================
:PythonCheck
call :CheckREG
if defined PYTHON goto :eof
echo 正在安裝 Python 3.14...
echo.
echo.
winget install Python.Python.3.14 --source winget --accept-package-agreements --accept-source-agreements
COLOR 0B
timeout /t 2 >nul
echo.
echo.
call :CheckREG
if defined PYTHON goto :eof
echo 安裝失敗 請手動安裝Python
pause
exit

::====================================================================================
:CheckREG
for %%R in (
    "HKCU\SOFTWARE\Python\PythonCore"
    "HKLM\SOFTWARE\Python\PythonCore"
    "HKLM\SOFTWARE\WOW6432Node\Python\PythonCore"
) do (
    for /f "tokens=2*" %%A in ('reg query %%R /s /v "ExecutablePath" 2^>nul ^| findstr "ExecutablePath"') do (
        if exist "%%B" (
            set "PYTHON=%%B"
            goto :eof
        )
    )
)
goto :eof

::====================================================================================
:GetAdmin
echo Set UAC = CreateObject^("Shell.Application"^) >"!temp!\getadmin.vbs"
echo UAC.ShellExecute "%~fs0", "", "", "runas", 1 >>"!temp!\getadmin.vbs"
"!temp!\getadmin.vbs"
del /f /q "!temp!\getadmin.vbs"
exit

::====================================================================================