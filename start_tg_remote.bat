@echo off
:: PyMacro TG Remote — Boot Launcher
:: Starts tg_remote.py silently in the background on Windows login.

cd /d "%~dp0"
start "" /B "C:\Program Files\Python310\python.exe" "%~dp0src\tg_remote.py"
