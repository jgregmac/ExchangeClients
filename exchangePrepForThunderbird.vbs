'exchangePrepForThunderbird.vbs script, J. Greg Mackinnon, 2015-10-22
' Kills any running Thunderbird processes, removes the legacy mailbox path prefix,
' sets the imap new mail check interval to 10 minutes, and restarts Thunderbird if 
' it was running.
' A backup copy of the userpref.js file is created when the script is run.
' If run silently, Thunderbird will be closed automatically and no message boxes will be displayed.
'
' Usage:
'  cscript.exe exchangePrepForThunderbird.vbs [/silent:(True|False)]
'
'Provides:
' RC=010 - Invalid input paramater.
' RC=101 - Could not Locate a Thunderbird user profiles storage directory.
' RC=102 - Could not locate a Thunderbird prefs.js file to modify.
' RC=200 - User chose not to terminate Thunderbird.  Script execution canceled.
' RC=201 - Could not stop the running Thunderbird processes.
' RC=301 - Could not write out changes to the prefs.js file.
' RC=401 - Could not determine the system architecture (used to locate Thunderbird.exe).
' RC=402 - Could not determine the path to Thunderbird.exe

Option Explicit

Const quote = """"
Const ForReading = 1
Const ForWriting = 2
Const ActivateAndDisplay = 1
Const noWaitOnDisplay = 0

'Declare Variables:
Dim aKills(1)
Dim bIsRunning, bRestore, bSilent
Dim cArgs
Dim dNow
Dim oShell, oFS, oFile, oLog
Dim re1, re2 'Regular Expressions
Dim sBadArg, sCmd, sErr, sKill, sLine, sLog, sLogRoot, sMsgTitle, sNewContents, sScrArg, sSilent, sTemp

'Set initial values:
sMsgTitle = "UVM Exchange Preparation Tool for Thunderbird"
aKills(0) = "Thunderbird.exe"
bRestore = False

'Instantiate Global Objects:
Set cArgs = WScript.Arguments.Named
Set oShell = CreateObject("WScript.Shell")
Set oFS  = CreateObject("Scripting.FileSystemObject")
Set re1 = New RegExp
Set re2 = New RegExp

'Initialize Regular Expression object to search for the Mailbox path prefix:
re1.Pattern    = "^user_pref\(""mail\.server\.server[2-9]\.server_sub_directory"""
re1.IgnoreCase = False
re1.Global     = False
'Initialize Regular Expression object to search for the IMAP retry interval:
'mail.server.server2.check_time
re2.Pattern    = "^user_pref\(""mail\.server\.server[2-9]\.check_time"""
re2.IgnoreCase = False
re2.Global     = False

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialize Logging
sLogRoot = "exchangePrepForThuderbird"
sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
dNow = Replace(Date(),"/","-") & "-" & Replace(Replace(Time()," ",""),":","_")
sLog = sLogRoot & "-" & dNow & ".log"
Set oLog = oFS.OpenTextFile(sTemp & "\" & sLog, 2, True)
' End Initialize Logging
'''''''''''''''''''''''''''''''''''''''''''''''''''

' Determine if we should run silently:
If cArgs.Exists("silent") Then
    sSilent = UCase(cArgs.Item("silent"))
echoAndLog "Silent argument provided with value: " & sSilent
    If sSilent = "TRUE" Then
        bSilent = True
    ElseIf sSilent = "FALSE" Then
        bSilent = False
    Else 
        echoAndLog "Silent argument is invalid.  Please enter either 'True' or 'False'."
        WScript.Quit
    End If
Else
    bSilent = False
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Define Functions
'
Sub subHelp
	echoAndLog "exchangePrepForThunderbird.vbs Script"
	echoAndLog "by J. Greg Mackinnon, University of Vermont"
	echoAndLog ""
	echoAndLog "Kills any running Thunderbird processes, removes the legacy mailbox path prefix, "
    echoAndLog "sets the imap new mail check interval to 10 minutes, and restarts Thunderbird "
	echoAndLog "if it was running."
    echoAndLog ""
	echoAndLog "A backup copy of the userpref.js file is created when the script is run."
    echoAndLog ""
    echoAndLog "If run silently, Thunderbird will be closed automatically and no message boxes will"
    echoAndLog "be displayed."
    echoAndLog ""
	echoAndLog "Returns:"
    echoAndLog " RC=010 - Invalid input paramater."
    echoAdnLog " RC=101 - Could not Locate a Thunderbird user profiles storage directory."
    echoAndLog " RC=102 - Could not locate a Thunderbird prefs.js file to modify."
    echoAndLog " RC=200 - User chose not to terminate Thunderbird.  Script execution canceled."
    echoAndLog " RC=201 - Could not stop the running Thunderbird processes."
    echoAndLog " RC=301 - Could not write out changes to the prefs.js file."
    echoAdnLog " RC=401 - Could not determine the system architecture (used to locate Thunderbird.exe)."
    echoAndLog " RC=402 - Could not determine the path to Thunderbird.exe."
    echoAndLog ""
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
    Dim iResponse
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
    echoAndLog "--------------------------------------------------"
    echoAndLog "Checking for processes to terminate..."
    'Cycle through found problematic processes and kill them.
    For Each oProc in cProcs
        echoAndLog "Found process " & oProc.Name & "."
        sErr = oProc.Name & " is currently running.  Click 'Okay' to exit " _
        & oProc.Name & " and continue with updates."
        iResponse = MsgBox(sErr, 33, sMsgTitle)
        if iResponse = 1 or bSilent then
            'Set this to look for errors that aren't fatal when killing processes.
            On Error Resume Next
            oProc.Terminate()
        else 
            sErr = "Task canceled, no changes made to Thunderbird.  Please start " _
            & "this tool again when Thunderbird is not running."
            msgBox sErr, 16, sMsgTitle
            WScript.Quit(200)
        end if
        Select Case Err.Number
            Case 0
                echoAndLog "Killed process " & oProc.Name & "."
                Err.Clear
                bKilled = True
            Case -2147217406
                echoAndLog "Process " & oProc.Name & " already closed."
                Err.Clear
            Case Else
                sErr = "Could not stop Thunderbird.  Please exit Thunderbird before running this script again."
                echoAndLog sErr
                echoAndLog "Error Number: " & Err.Number
                echoAndLog "Error Description: " & Err.Description
                echoAndLog "Finished process termination function with error."
                echoAndLog "--------------------------------------------------"
                echoAndLog vbCrLf & "script finished."
                echoAndLog "************************************************************" & vbCrLf
                If bSilent = False Then
                    msgBox sErr, 16, sMsgTitle
                End If
                WScript.Quit(201)
        End Select
    Next
    'Resume normal error handling.
    On Error Goto 0
    echoAndLog "Finished process termination function."
    echoAndLog "--------------------------------------------------"
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
On Error Resume Next
Set oProfileDir = oFS.GetFolder(sUProfile & "\AppData\Roaming\Thunderbird\Profiles")
Select Case Err.Number
    Case 0
        echoAndLog "Located Thunderbird Profiles directory."
        Err.Clear
    Case Else
        sErr = "Could not find a Thunderbird user profiles storage directory! Aborting Script!"
        echoAndLog sErr
        echoAndLog "*----------------------------------------------------------*"
        echoAndLog "************************************************************"
        If bSilent = False Then
            msgBox sErr, 16, sMsgTitle
        End If
        WScript.Quit(101)
End Select
On Error Goto 0
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
        echoAndLog aPrefs(iSize)
        iSize = iSize + 1
    End If
Next
If bPrefsFound Then
    echoAndLog "Count of preference files located: " & iSize
Else
    sErr = "No Thunderbird users preferences files located.  Quitting..."
    echoAndLog sErr
    echoAndLog "*----------------------------------------------------------*"
    echoAndLog "************************************************************"
    If bSilent = False Then
        msgBox sErr, 16, sMsgTitle
    End If
    WScript.Quit(102)
End If
echoAndLog "*----------------------------------------------------------*"


echoAndLog vbCrLf & "*----------------------------------------------------------*"
echoAndLog "Begin mailbox path prefix remediation:" & vbCrLf
Dim bMatch1, bMatch2
bMatch1 = False
bMatch2 = False
Dim bWritePrefs
bWritePrefs = False
Dim i
'Create a loop to process all prefs.js files that were located:
For i = LBound(aPrefs) To UBound(aPrefs)
    'Test each line prefs.js for line defining the mailbox path prefix, save any non-matching line to sNewContents: 
    sNewContents = "" 'Clear an existing content from sNewContents (from previous pass through this loop):
    echoAndLog "Reading Pref.js contents: " & vbCrLf & aPrefs(i)
    Set oFile = oFS.OpenTextFile(aPrefs(i), ForReading)
    Do Until oFile.AtEndOfStream
        sLine = oFile.ReadLine
        bMatch1 = re1.Test(sLine)
        bMatch2 = re2.Test(sLine)
        If bMatch1 Then
            echoAndLog "Found the mailbox path prefix in prefs.js:"
            echoAndLog sLine
            'Flag to write changes to prefs.js to file:
            bWritePrefs = True
        ElseIf bMatch2 Then
            echoAndLog "Found imap retry interval in prefs.js:"
            echoAndLog sLine
            'Now modify sLine with the preferred retry interval...
            'First, locate the current retry value, and capture to "iVal"
            Dim iLen, iPos, iVal 
            iPos = InStr(sLine, ",")
            iLen = (InStr(sLine, ")") - iPos) - 1
            iVal = Trim(Mid(sLine, (iPos + 1), iLen))
            'Change the value if greater than 10:
            if iVal > 10 then
                echoAndLog "Server retry value is greater than 10"
                sLine = Replace(sLine,iVal,"10")
                echoAndLog "New preference file entry: " & vbCrlf & sLine
                'Write out the modified line to "sNewContents"
                sNewContents = sNewContents & sLine & vbCrLf
                'Flag to write changes to prefs.js to file:
                bWritePrefs = True
            end if
        Else
            'Write out the unmodified line to sNewContents
            sNewContents = sNewContents & sLine & vbCrLf
        End If
    Loop
    oFile.Close
    ' If we found a match, kill Thunderbird and write the changes out to file:
    If bWritePrefs Then
        'Determine if Thunderbird is running and kill it 
        ' (script will exit if fKillProcs cannot terminate Thunderbird):
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
                sErr = "Could not write out changes to prefs.js! Aborting Script!"
                echoAndLog sErr
                If bSilent = False Then
                    msgBox sErr, 16, sMsgTitle
                End If
                WScript.Quit(301)
        End Select
        On Error Goto 0
    Else
        echoAndLog "No problematic settings found in this prefs.js file."
    End If
Next

If bWritePrefs = False Then
    sErr = "No problematic settings were found in the any user preferences files." & vbCrLf _
    & "Thunderbird does not need to be updated."
    echoAndLog sErr
    echoAndLog "*----------------------------------------------------------*"
    echoAndLog "************************************************************"
    If bSilent = False Then
        msgBox sErr, 0, sMsgTitle
    End If
    WScript.Quit(0)
End If

echoAndLog "End mailbox path prefix remediation."
echoAndLog "*----------------------------------------------------------*"

If bIsRunning Then
    echoAndLog vbCrLf & "*----------------------------------------------------------*"
    echoAndLog "Thunderbird was previously running.  Restarting Thunderbird..." & vbCrLf
    echoAndLog "--------------------------------------------------"
    echoAndLog "Determining system architecture..."
    Dim oWMISvc
    Set oWMISvc = GetObject("winmgmts:{impersonationLevel=impersonate, (Debug)}\\.\root\cimv2")
    Dim sQuery 
    sQuery = "Select OSArchitecture from Win32_OperatingSystem"
    Dim cArch
    On Error Resume Next
    Set cArch = oWMISvc.ExecQuery(sQuery, "WQL", 48)
    Select Case Err.Number
        Case 0
            echoAndLog "System architecture queried successfully."
            Err.Clear
        Case Else
            sErr = "Could not determine the system architecture." & vbCrLf _
                & "Thunderbird preferences were updated, but you will need to restart Thunderbird manually."
            echoAndLog sErr
            echoAndLog "*----------------------------------------------------------*"
            echoAndLog "************************************************************"
            If bSilent = False Then
                msgBox sErr, 32, sMsgTitle
            End If
            WScript.Quit(401)
    End Select
    On Error Goto 0
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
    Dim sTPath
    sTPath = sProgramFiles & "\Mozilla Thunderbird\thunderbird.exe"
    If oFS.FileExists(sTPath) Then
        echoAndLog "Found Thunderbird.exe in path:"
        echoAndLog sTPath
    Else 
        sErr = "Could not find the Thunderbird program files." & vbCrLf & vbCrLf _
            & "Thunderbird preferences were updated, but you will need to restart Thunderbird manually."
        echoAndLog sErr
        echoAndLog "*----------------------------------------------------------*"
        echoAndLog "************************************************************"
        If bSilent = False Then
            msgBox sErr, 32, sMsgTitle
        End If
        WScript.Quit(402)
    End If
    echoAndLog "--------------------------------------------------"
    'Run the executable in sTPath.  Ensure that it will not terminate when this script exits.
    echoAndLog "Now restarting Thunderbird..."
    ' The "ActivateAndDisplay" preference does not really matter. All flags do the same thing.
    'oShell.Run quote & sTPath & quote,ActivateAndDisplay,noWaitOnDisplay
    echoAndLog "*----------------------------------------------------------*"
    echoAndLog "************************************************************"
End If

sErr = "Thunderbird preferences files updated successfully." & vbCrLf & "Have a happy migration to Exchange."
echoAndLog sErr
If bSilent = False Then
    msgBox sErr, 0, sMsgTitle
End If
oLog.Close
wscript.quit(0)