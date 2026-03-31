@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0audio-switcher.ps1" -ConfigPath "%~dp0config.json"
