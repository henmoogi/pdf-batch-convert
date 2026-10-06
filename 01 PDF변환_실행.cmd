@echo off
chcp 65001 >nul
rem PDF batch converter launcher. Uses Windows built-in PowerShell (no Python needed).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0automation\PDF변환_시작.ps1"
