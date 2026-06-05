@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0start.ps1" %*
