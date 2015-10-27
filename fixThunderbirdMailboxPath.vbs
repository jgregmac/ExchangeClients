'fixThuderbirdMailPrefix.vbs script, J. Greg Mackinnon, 2015-10-22
' Kills any running Thunderbird processes, removes the legacy mailbox path prefix, 
' and restarts Thunderbird if it was running.
' A backup copy of the userpref.js file is created when the script is run.  This file can be 
' restored by running this script with the "restore" switch.
'Provides:
' RC=101 - Error terminating the requests processes
' RC=100 - Invalid input parameters
' Other return codes - Pass-though of return code from WShell.Exec.Run using the provided input parameters

Option Explicit

Const quote = """"
Const ForReading = 1
Const ForWriting = 2
Const ActivateAndDisplay = 1
Const noWaitOnDisplay = 0

'Declare Variables:
Dim aKills(1)
Dim bIsRunning, bMatch, bPathPrefixExists, bRestore
Dim cScrArgs
Dim iReturn
Dim oShell, oFS, oFile, oLog
Dim re
Dim sBadArg, sCmd, sKill, sLine, sLog, sNewContents, sScrArg, sTemp

'Set initial values:
aKills(0) = "thunderbird.exe"
bRestore = False
bMatch = False
bPathPrefixExists = False
iReturn = 0

'Instantiate Global Objects:
Set oShell = CreateObject("WScript.Shell")
Set oFS  = CreateObject("Scripting.FileSystemObject")
Set re = New RegExp

'Initialize Regular Expression object to search for the Mailbox path prefix:
re.Pattern    = "^user_pref\(""mail\.server\.server[2-9]\.server_sub_directory"""
re.IgnoreCase = False
re.Global     = False

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialize Logging
sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
sLog = "fixThuderbirdMailPrefix.log"
Set oLog = oFS.OpenTextFile(sTemp & "\" & sLog, 2, True)
' End Initialize Logging
'''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Define Functions
'
Sub subHelp
	echoAndLog "KillAndExec.vbs Script"
	echoAndLog "by J. Greg Mackinnon, University of Vermont"
	echoAndLog ""
	echoAndLog "Kills named processes and runs the provided executable."
	echoAndLog "Logs output to 'KillAndExec-[exeName].log' in the %temp% directory."
	echoAndLog ""
	echoAndLog "Required arguments and syntax:"
	echoAndLog "/kill:""[process1];[process2]..."""
	echoAndLog "     Specify the image name of one or more processes to terminate."
	echoAndLog "/exe:""[ExecutableFile.exe]"""
	echoAndLog "     Specify the name of the executable to run."
	echoAndLog ""
	echoAndLog "Optional arguments:"
	echoAndLog "/args""[arg1];[arg2];[arg3]..."""
	echoAndLog "     Specify one or more arguments to pass to the executable."
End Sub

function echoAndLog(sText)
'EchoAndLog Function:
' Writes string data provided by "sText" to the console and to Log file
' Requires: 
'     sText - a string containing text to write
'     oLog - a pre-existing Scripting.FileSystemObject.OpenTextFile object
	'If we are in cscript, then echo output to the command line:
	If LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then
		wscript.echo sText
	end if
	'Write output to log either way:
	oLog.writeLine sText
end function

function fKillProcs(aKills)
' Requires:
'     aKills - an array of strings, with each entry being the name of a running process.   
	Dim bKilled
    Dim cProcs
	Dim sProc, sQuery
	Dim oWMISvc, oProc
    
    bKilled = False
	Set oWMISvc = GetObject("winmgmts:{impersonationLevel=impersonate, (Debug)}\\.\root\cimv2")
	sQuery = "Select Name from Win32_Process Where " 'Root query, will be expanded.	
	'Complete the query string using process names in "aKill"
	for each sProc in aKills
		sQuery = sQuery & "Name = '" & sProc & "' OR "
	next
	'Remove the trailing " OR" from the query string
	sQuery = Left(sQuery,Len(sQuery)-3)

	'Create a collection of processes named in the constructed WQL query
	Set cProcs = oWMISvc.ExecQuery(sQuery, "WQL", 48)
	echoAndLog vbCrLf & "--------------------------------------------------"
	echoAndLog "Checking for processes to terminate..."
    'Set this to look for errors that aren't fatal when killing processes.
    On Error Resume Next
    'Cycle through found problematic processes and kill them.
    For Each oProc in cProcs
       echoAndLog "Found process " & oProc.Name & "."
       oProc.Terminate()
       Select Case Err.Number
           Case 0
               echoAndLog "Killed process " & oProc.Name & "."
               Err.Clear
               bKilled = True
           Case -2147217406
               echoAndLog "Process " & oProc.Name & " already closed."
               Err.Clear
           Case Else
               echoAndLog "Could not kill process " & oProc.Name & "! Aborting Script!"
               echoAndLog "Error Number: " & Err.Number
               echoAndLog "Error Description: " & Err.Description
               echoAndLog "Finished process termination function with error."
               echoAndLog "--------------------------------------------------"
               echoAndLog vbCrLf & "script finished."
               echoAndLog "************************************************************" & vbCrLf
               WScript.Quit(201)
       End Select
    Next
	'Resume normal error handling.
	On Error Goto 0
	echoAndLog "Finished process termination function."
	echoAndLog "------------------------------------------------------------"
    If bKilled Then
        fKillProcs = True
    Else
        fKillProcs = False
    End If
end function
'
' End Define Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''

echoAndLog vbCrLf & "************************************************************"
echoAndLog "*----------------------------------------------------------*"
echoAndLog "Locating Thunderbird Prefs.js..." & vbCrLf
Dim sUProfile
sUProfile = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
Dim oProfileDir
Set oProfileDir = oFS.GetFolder(sUProfile & "\AppData\Roaming\Thunderbird\Profiles")
Dim bPrefsFound
bPrefsFound = False
Dim cProfiles
Set cProfiles = oProfileDir.subFolders
Dim iSize
iSize = CInt(0)
ReDim aPrefs(iSize)
Dim oProfile
For Each oProfile in cProfiles
    Dim sPrefsPath
    sPrefsPath = oProfile.Path & "\prefs.js"
    Dim bExists
    bExists = False
    bExists = oFS.FileExists(sPrefsPath)
    If bExists Then
        bPrefsFound = True
        ReDim Preserve aPrefs(iSize)
        aPrefs(iSize) = sPrefsPath
        echoAndLog "Prefs.js exists in path: "
        echoAndLog vbCrLf & aPrefs(iSize)
        iSize = iSize + 1
    End If
Next
If bPrefsFound Then
    echoAndLog vbCrLf & "Count of preference files located: " & iSize
Else
    echoAndLog vbCrLf & "No prefs.js files located.  Quitting..."
    WScript.Quit(201)
End If
echoAndLog "*----------------------------------------------------------*"


echoAndLog vbCrLf & "*----------------------------------------------------------*"
echoAndLog "Begin mailbox path prefix remediation:"
Dim i
For i = LBound(aPrefs) To UBound(aPrefs)
    'Test each line prefs.js for line defining the mailbox path prefix, save any non-matching line to sNewContents: 
    wscript.echo "Pref contents: " & aPrefs(i)
    Set oFile = oFS.OpenTextFile(aPrefs(i), ForReading)
    Do Until oFile.AtEndOfStream
        sLine = oFile.ReadLine
        bMatch = re.Test(sLine)
        If bMatch Then
            echoAndLog "Found the mailbox path prefix in prefs.js."
            bPathPrefixExists = True
        Else
            sNewContents = sNewContents & sLine & vbCrLf
        End If
    Loop
    oFile.Close
    ' If we found a match, kill Thunderbird and write the changes out to file:
    If bPathPrefixExists Then
        'Determine if Thunderbird is running and kill it:
        bIsRunning = fKillProcs(aKills)
        echoAndLog "Now updatating the contents of prefs.js, excluding the mailbox path prefix."
        Set oFile = oFS.OpenTextFile(aPrefs(i), ForWriting)
        oFile.Write sNewContents
        On Error Resume Next
        oFile.Close
        Select Case Err.Number
            Case 0
                echoAndLog "Successfully updated prefs.js."
                Err.Clear
            Case Else
                WScript.Quit(301)
        End Select
        On Error Goto 0
    Else
        echoAndLog "Mailbox path prefix not found in prefs.js."
    End If
Next
echoAndLog "End mailbox path prefix remediation."
echoAndLog "*----------------------------------------------------------*"

If bIsRunning Then
    echoAndLog vbCrLf & "*----------------------------------------------------------*"
    echoAndLog "Thunderbird was previously running.  Restarting Thunderbird..."
    echoAndLog "--------------------------------------------------"
    echoAndLog "Determining system architecture..."
    'Determine System Architecture:
    Dim oWMISvc
    Set oWMISvc = GetObject("winmgmts:{impersonationLevel=impersonate, (Debug)}\\.\root\cimv2")
    Dim sQuery 
    sQuery = "Select OSArchitecture from Win32_OperatingSystem"
    Dim cArch
    Set cArch = oWMISvc.ExecQuery(sQuery, "WQL", 48)
    Dim bIs64 
    bIs64 = False
    Dim oArch
    Dim sArch
    For Each oArch in cArch
        sArch = CStr(oArch.OSArchitecture)
        If InStr(sArch,"64-bit") > 0 Then
            bIs64 = True
            echoAndLog "System is 64-bit."
        Else
            echoAndLog "System is 32-bit."
        End If
    Next
    echoAndLog "--------------------------------------------------"
    echoAndLog "Determining path to Thunderbird Program Files..."
    'Determine Program Files path:
    Dim sProgramFiles
    If bIs64 Then
        sProgramFiles = oShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")
    Else
        sProgramFiles = oShell.ExpandEnvironmentStrings("%ProgramFiles%")
    End If 
    echoAndLog "Program Files directory path is: " 
    echoAndLog sProgramFiles
    'Determine Path to Thunderbird.exe
    Dim sTPath
    sTPath = sProgramFiles & "\Mozilla Thunderbird\thunderbird.exe"
    If oFS.FileExists(sTPath) Then
        echoAndLog "Found Thunderbird.exe in path:"
        echoAndLog sTPath
    Else 
        echoAndLog "Could not find Thunderbird.exe"
    End If
    echoAndLog "--------------------------------------------------"
    'Run the executable in sTPath.  Ensure that it will not terminate when this script exits.
    echoAndLog "Now restarting Thunderbird..."
    oShell.Run quote & sTPath & quote,ActivateAndDisplay,noWaitOnDisplay
    echoAndLog "*----------------------------------------------------------*"
    echoAndLog "************************************************************"
End If


oLog.Close
wscript.quit(iReturn)