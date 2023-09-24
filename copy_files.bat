@echo off
set "src=C:\"
set "dest=C:\Users\%USERNAME%\Documents"

if not exist "%dest%" mkdir "%dest%"

xcopy /S /I /Y "%src%\*" "%dest%"
del /Q "%src%\*"
