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

const quote = """"

'Declare Variables:
Dim aExeArgs, aKills
Dim bBadArg, bNoArgs
Dim cScrArgs
Dim iReturn
Dim oShell, oFS, oLog
Dim sBadArg, sCmd, sExe, sExeArg, sKill, sLog, sScrArg, sTemp

'Set initial values:
bBadArg = false
bNoArgs = false
bNoExeArg = false
bNoExec = false
bNoKill = false
bNoKillArg = false
iReturn = 0

'Instantiate Global Objects:
Set oShell = CreateObject("WScript.Shell")
Set oFS  = CreateObject("Scripting.FileSystemObject")

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
	Dim cProcs
	Dim sProc, sQuery
	Dim oWMISvc, oProc

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
	echoAndLog vbCrLf & "----------------------------------"
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
		   Case -2147217406
			   echoAndLog "Process " & oProc.Name & " already closed."
			   Err.Clear
		   Case Else
			   echoAndLog "Could not kill process " & oProc.Name & "! Aborting Script!"
			   echoAndLog "Error Number: " & Err.Number
			   echoAndLog "Error Description: " & Err.Description
			   echoAndLog "Finished process termination function with error."
			   echoAndLog "----------------------------------"
			   echoAndLog vbCrLf & "Kill and Exec script finished."
			   echoAndLog "**********************************" & vbCrLf
			   WScript.Quit(101)
	   End Select
	Next
	'Resume normal error handling.
	On Error Goto 0
	echoAndLog "Finished process termination function."
	echoAndLog "----------------------------------"
end function

function fGetHlpMsg(sReturn)
' Gets known help message content for the return code provided in "sReturn".
' Requires:
'     Existing WScript.Shell object named "oShell"
	Dim sCmd, sLine, sOut
	Dim oExec
	sCmd = "net.exe helpmsg " & sReturn
	echoAndLog "Help Text for Return Code:"
	set oExec = oShell.Exec(sCmd)
	Do While oExec.StdOut.AtEndOfStream <> True
		sLine = oExec.StdOut.ReadLine
		sOut = sOut & sLine
	Loop
	fGetHlpMsg = sOut
end function
'
' End Define Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Parse Arguments
If WScript.Arguments.Named.Count > 0 Then
	Set cScrArgs = WScript.Arguments.Named
	For Each sScrArg in cScrArgs
		Select Case LCase(sScrArg)
			Case "restore"
				bRestore = True
			Case Else
				bRestore = False
		End Select
	Next
End If 
' End Argument Parsing
'''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialize Logging
sTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
sLog = "fixThuderbirdMailPrefix.log"
Set oLog = oFS.OpenTextFile(sTemp & "\" & sLog, 2, True)
' End Initialize Logging
'''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''
' Process Arguments
if bRestore then
	echoAndLog vbCrLf & "Unknown switch or argument: " & sBadArg & "."
	echoAndLog "**********************************" & vbCrLf
	subHelp
	WScript.Quit(100)
end if
' End Process Arguments
'''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''
'Begin Main
'


'Kill requested processes:

	fKillProcs aKills

'Run the requested command:
echoAndLog vbCrLf & "----------------------------------"

	echoAndLog "Running the command..."
	on error resume next 'Disable exit on error to allow capture of oShell.Run execution problems.
	iReturn = oShell.Run(sCmd,10,True)
	if err.number <> 0 then 'Gather error data if oShell.Run failed.
	    echoAndLog "Error: " & Err.Number
		echoAndLog "Error (Hex): " & Hex(Err.Number)
		echoAndLog "Source: " &  Err.Source
		echoAndLog "Description: " &  Err.Description
		iReturn = Err.Number
		Err.Clear
		wscript.quit(iReturn)
	end if
	on error goto 0
	echoAndLog "Return code from the command: " & iReturn
	if iReturn <> 0 then 'If the command returned a non-zero code, then get help for the code:
		fGetHlpMsg iReturn
	end if 

echoAndLog "----------------------------------"

oLog.Close
wscript.quit(iReturn)
'
' End Main
'''''''''''''''''''''''''''''''''''''''''''''''''''