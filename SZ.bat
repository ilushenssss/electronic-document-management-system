@echo off
rem Sluzhebnye zapiski: zapusk bez ustanovki Python (vsyo v papke environment)
start "" "%~dp0environment\python\pythonw.exe" "%~dp0main.pyw" %*
