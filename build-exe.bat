@echo off
set CSC=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
set ICON=%~dp0assets\app-icon.ico
if not exist "%CSC%" (
  echo C# compiler not found: %CSC%
  exit /b 1
)

if exist "%ICON%" (
  "%CSC%" /nologo /target:winexe /win32icon:"%ICON%" /out:"%~dp0AudioSwitcher.exe" /reference:System.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Web.Extensions.dll "%~dp0AudioSwitcher.cs"
) else (
  "%CSC%" /nologo /target:winexe /out:"%~dp0AudioSwitcher.exe" /reference:System.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:System.Web.Extensions.dll "%~dp0AudioSwitcher.cs"
)
