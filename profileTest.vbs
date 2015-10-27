Set oShell = CreateObject("WScript.Shell")
Set oFS  = CreateObject("Scripting.FileSystemObject")
sUProfile = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
Set oTBirdProfileDir = oFS.GetFolder(sUProfile & "\AppData\Roaming\Thunderbird\Profiles")
Wscript.echo oTBirdProfileDir.Path
Set cTBirdProfiles = oTBirdProfileDir.subFolders
iAPrefsSize = CInt(1)
ReDim aPrefs(iAPrefsSize)
For Each oProfile in cTBirdProfiles
    wScript.Echo oProfile.Path
    sPrefsJSPath = oProfile.Path & "\prefs.js"
    bPrefsExist = oFS.FileExists(sPrefsJSPath)
    If bPrefsExist Then
        ReDim Preserve aPrefs(iAPrefsSize)
        wscript.echo "Prefs.js exists!  Let's continue..."
        aPrefs(iAPrefsSize) = sPrefsJSPath
        WScript.echo aPrefs(iAPrefsSize)
        iAPrefsSize = iAPrefsSize + 1
    Else
        wscript.echo "Prefs.js is not present."
        'Wscript.quit(201)
    End If
Next
'sPrefsJSPath = sTBirdProfile.Item & "\prefs.js"
'WScript.Echo sPrefsJSPath