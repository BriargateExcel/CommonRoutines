VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorFileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Purpose
' This class implements an error handling approach where a caller can
' write lines of text to a .txt file.
' When the class terminates, it gives the user a MsgBox stating that
' errors were found, an error file was created, and where to find
' the error file.
' If there were no errors reported to this instance of this class,
' there is no messsage to the user and no error files are created.
' Only when an error is reported to this clas, is an error file created
' and only when there is an error file does the user get a message
' about the error file.
' If there are no errors reported, this class remains silent.
'
' Usage
' Create an instance of the class with statements like:
'   Dim MyErrorFile as ErrorFileClass
'   Set MyErrorFile = new ErrorFileClass
'
' Report an error by calling the WriteErrorLine routine:
'   MyErrorFile.WriteErrorLine "This is an error message"
'
' Error messages can span multiple lines of text:
'   MyErrorFile.WriteErrorLine "Line 1" & vbCrLf _
'       "Line 2" & vbCrLf _
'       "Line 3"
'
' Write a blank line to the error file with:
'   MyErrorFile.WriteBlankErrorLines
' To write multiple blank lines, provide a number.
'
' When MyErrorFile goes out of scope, the Class_Terminate
' routine fires and reports the existence and location of the error
' file, if an error was reported.
' If no error was reported, there is no error file and no user message.

Private Const Module_Name As String = "ErrorFileClass."

Private Const ErrorMessageName As String = "Error Messages"
Private pErrorStream As TextStream

Public Sub WriteErrorLine(ByVal ErrorMessage As String)

    ' This routine writes one line to pErrorStream
    
    Const RoutineName As String = Module_Name & "WriteErrorLine"
    On Error GoTo ErrorHandler
    
    #If DebugOn Then
        Dim cPerf As PerformanceClass
        If gbDebug(RoutineName) Then
            Set cPerf = New PerformanceClass
            cPerf.SetRoutine RoutineName
        End If
    #End If
    
    If pErrorStream Is Nothing Then CreateErrorStream
    
    pErrorStream.WriteLine ErrorMessage
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WriteErrorLine

Public Sub WriteBlankErrorLines(Optional ByVal NumLines As Long = 1)

    ' This routine writes NumLines blank lines to pErrorStream
    
    Const RoutineName As String = Module_Name & "WriteBlankErrorLine"
    On Error GoTo ErrorHandler
    
    #If DebugOn Then
        Dim cPerf As PerformanceClass
        If gbDebug(RoutineName) Then
            Set cPerf = New PerformanceClass
            cPerf.SetRoutine RoutineName
        End If
    #End If
    
    If pErrorStream Is Nothing Then CreateErrorStream
    
    pErrorStream.WriteBlankLines NumLines
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WriteBlankErrorLine

Private Sub CreateErrorStream()

    ' This routine creates an Error Message folder
    ' on the user's desktop if one does not already exist.
    ' It then finds the next available Error Message File number and opens it
    
    Const RoutineName As String = Module_Name & "WriteErrorLine"
    On Error GoTo ErrorHandler
    
    #If DebugOn Then
        Dim cPerf As PerformanceClass
        If gbDebug(RoutineName) Then
            Set cPerf = New PerformanceClass
            cPerf.SetRoutine RoutineName
        End If
    #End If
    
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    Dim DF As String
    DF = DesktopFolder
    
    Dim ErrorMessageFolderPath As String
    ErrorMessageFolderPath = DF & Application.PathSeparator & ErrorMessageName
    
    If FSO.FolderExists(ErrorMessageFolderPath) Then
        ' Folder exists. Do nothing.
    Else
        ' Folder does not exist. Create it.
        FSO.CreateFolder ErrorMessageFolderPath
    End If
    ' Error message folder now exists on Desktop
 
    Dim FileNumber As Long
    FileNumber = 1
    
    
    While FSO.FileExists(ErrorMessageFolderPath & Application.PathSeparator & ErrorMessageName & FileNumber & ".txt")
        FileNumber = FileNumber + 1
    Wend
 
    Dim ErrorMessageFile As String
    ErrorMessageFile = ErrorMessageFolderPath & Application.PathSeparator & ErrorMessageName & FileNumber & ".txt"
    
    Set pErrorStream = FSO.CreateTextFile(ErrorMessageFile)
 
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' WriteErrorLine

Private Sub Class_Terminate()
    ' Tell the user if an error file was created
    If Not pErrorStream Is Nothing Then
        MsgBox "Check the Error Messages folder on your Desktop." & vbCrLf _
             & "See the Error Messages file in that folder " & _
               "for a list of the error messages found.", _
               vbOKOnly Or vbExclamation, _
               "Errors Found"
        pErrorStream.Close
        Set pErrorStream = Nothing
    Else
        '        MsgBox "No errors found", vbOKOnly Or vbExclamation, "No Errors Found"
    End If

End Sub                                          ' Class_Terminate


