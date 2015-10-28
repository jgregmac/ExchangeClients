@echo off
REM @echo on

set fname=fixThunderbirdMailboxPath
echo Fname is %fname%
pause
Echo Cleaning up old builds...
del %fname%.exe
del %fname%.7z

Echo Packaing the script...
..\bin\7zr a .\%fname%.7z .\%fname%.vbs
copy /b ..\bin\7zS.sfx + .\config.txt + .\%fname%.7z .\%fname%.exe
mt.exe -nologo -manifest ".\%fname%.manifest" -outputresource:".\%fname%.exe"
Echo Cleaning up build environment...
del /f /q .\%fname%.7z