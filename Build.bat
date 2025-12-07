@echo off
setlocal

rem === Get the directory of this script (always ends with \) ===
set SCRIPT_DIR=%~dp0

echo Building SwitchPowerTray from "%SCRIPT_DIR%"

rem === Try to locate csc.exe from .NET Framework 4 ===
set CSC_EXE=

if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" (
    set "CSC_EXE=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
) else if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe" (
    set "CSC_EXE=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)

if "%CSC_EXE%"=="" (
    echo.
    echo ERROR: Could not find csc.exe from .NET Framework v4.0.30319.
    echo Make sure ".NET Framework 4.x Developer Pack" or "Build tools" are installed,
    echo or run this script from a "Developer Command Prompt for Visual Studio".
    echo.
    pause
    exit /b 1
)

echo Using compiler: %CSC_EXE%
echo.

"%CSC_EXE%" ^
  /nologo /target:winexe /platform:anycpu ^
  /win32icon:"%SCRIPT_DIR%appicon.ico" ^
  /out:"%SCRIPT_DIR%SwitchPowerTray.exe" ^
  /r:System.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll ^
  /resource:"%SCRIPT_DIR%Desktop_Dark.ico","Icon.Desktop.Dark.ico" ^
  /resource:"%SCRIPT_DIR%Desktop_Light.ico","Icon.Desktop.Light.ico" ^
  /resource:"%SCRIPT_DIR%Laptop_Dark.ico","Icon.Laptop.Dark.ico" ^
  /resource:"%SCRIPT_DIR%Laptop_Light.ico","Icon.Laptop.Light.ico" ^
  /resource:"%SCRIPT_DIR%Bolt_Dark.ico","Icon.Bolt.Dark.ico" ^
  /resource:"%SCRIPT_DIR%Bolt_Light.ico","Icon.Bolt.Light.ico" ^
  /resource:"%SCRIPT_DIR%Moon_Dark.ico","Icon.Moon.Dark.ico" ^
  /resource:"%SCRIPT_DIR%Moon_Light.ico","Icon.Moon.Light.ico" ^
  "%SCRIPT_DIR%SwitchPowerTray.cs"

if errorlevel 1 (
    echo.
    echo Build FAILED.
    echo.
    pause
    exit /b 1
) else (
    echo.
    echo Build SUCCEEDED: "%SCRIPT_DIR%SwitchPowerTray.exe"
    echo.
    pause
)

endlocal
