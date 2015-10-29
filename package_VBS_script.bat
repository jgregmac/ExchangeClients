REM Thunderbird fix script packaging script.
REM Requires - 
REM   - 7zr.exe and 7zS.sfx files from 7-Zip Extras, in a sibling "bin" folder
REM   - A config.txt with instructions for what 7zS is supposed to do,
REM   - signtool.exe from the Windows SDK
REM   - A PKCS12 (PFX)-formatted code signing certificate (mine is in ..\etc\)
REM   - ResourceHacker.exe: 
REM       a free utility that (among other features) allows application manifest manipulation.
REM   - A script and manifest to package, both with root name specified in "fname".
REM  
@echo off

set fname=fixThunderbirdMailboxPath
set SDKPath="C:\Program Files (x86)\Windows Kits\10\bin\x86"
set TimeStampURL="http://timestamp.verisign.com/scripts/timstamp.dll"
set /P CertPath="Enter the full path to the PKCS12/PFX signing certificate:"
set /P CertPass="Enter the password for certificate file:"

Echo Cleaning up old builds...
del %fname%.exe
del %fname%.7z
Echo.
Echo Creating 7z archive:
..\bin\7zr a .\%fname%.7z .\%fname%.vbs
Echo.
Echo Appending 7-zip self-extractor:
copy /b ..\bin\7zS.sfx + .\config.txt + .\%fname%.7z .\%fname%.exe
echo %errorlevel%
Echo.
Echo Applying Manifest:
REM mt.exe is known to corrupt executables that did not come out of Visual Studio.  Use ResourceHacker.
REM %SDKPath%\mt.exe -nologo -manifest ".\%fname%.manifest" -outputresource:".\%fname%.exe;#1"
..\bin\resource_hacker\ResourceHacker.exe -addoverwrite %fname%.exe, %fname%.exe, %fname%.manifest, 24,1,
Echo Return code from ResourceHacker: %ERRORLEVEL%
Echo.
Echo Cleaning up build environment...
del /f /q .\%fname%.7z
Echo.
Echo You still need to sign this thing:
%SDKPath%\signtool.exe sign /f "%CertPath%" /p "%CertPass%" /t %TimeStampUrl% /v %fname%.exe