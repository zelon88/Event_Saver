'File Name: es.vbs
'Version: v1.0, 10/31/2019
'Author: Justin Grimes, 10/31/2019

'Adapted from...
'https://stackoverflow.com/questions/21738159/extracting-error-logs-from-windows-event-viewer
'----------------------------------------

Option Explicit
Dim LOG_FILE, strComputer, objWMIService, objFSO, objLogFile, objLogFile2, objEvent, colItems
'----------------------------------------

'----------------------------------------
'Declare global variables
LOG_FILE = "C:\ProgramData\es\es.dat"
strComputer = "."
'----------------------------------------

'----------------------------------------
'Declare objects.
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objFSO = CreateObject("Scripting.FileSystemObject")
'----------------------------------------

'----------------------------------------
'Create a log file if none exists.
If Not objFSO.FileExists(LOG_FILE) Then
  Set objLogFile2 = objFSO.CreateTextFile(LOG_FILE)
  objLogFile2.Close
End If
'----------------------------------------

'----------------------------------------
'Open the log file.
Set objLogFile = objFSO.OpenTextFile(LOG_FILE, 8)
'----------------------------------------

'----------------------------------------
'Declare a function for writing log data.
Function writeLog(strText)
  objLogFile.WriteLine(strText)
End Function
'----------------------------------------

'----------------------------------------
'Write some welcome text to the log file.
writeLog("Starting es on: " & Date & VBNewLine)
'----------------------------------------

'----------------------------------------
'Execute a WMI query against the local Security logs.
Set colItems = objWMIService.ExecQuery("Select * from Win32_NTLogEvent WHERE LogFile='Security'")
For Each objEvent In colItems
  On Error Resume Next
    If (objEvent.EventCode >= 1000 And objEvent.EventCode <= 1003) Then 
        writeLog "--------------------"
        writeLog "Event Code: " & objEvent.EventCode & VBNewLine
        writeLog "Time Generated: " & objEvent.TimeGenerated
        writeLog "Time Written: " & objEvent.TimeWritten
        writeLog "Event Identifier: " & objEvent.EventIdentifier        
        writeLog "Category: " & objEvent.Category
        writeLog "Category String: " & objEvent.CategoryString
        writeLog "Computer Name: " & objEvent.ComputerName
        writeLog "Data: " & objEvent.Data
        writeLog "Insertion Strings: " & objEvent.InsertionStrings
        writeLog "Logfile: " & objEvent.Logfile
        writeLog "Message: " & objEvent.Message
        writeLog "Record Number: " & objEvent.RecordNumber
        writeLog "Source Name: " & objEvent.SourceName
        writeLog "Type: " & objEvent.Type
        writeLog "User: " & objEvent.User 
        writeLog "--------------------" & VBNewLine  
    End If
Next
'----------------------------------------

'----------------------------------------
'Close the logfile.
objLogFile.Close
'----------------------------------------