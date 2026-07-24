@echo off
REM Chameleon (Portable) launcher - double-click to run from a USB stick.
REM Runs the watcher hidden with no window. Close it from Task Manager
REM (look for powershell.exe) or log off to stop it.
start "" /min powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Chameleon-Portable.ps1"
