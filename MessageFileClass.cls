VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MessageFileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Changes
' 9/15/19
'       Changed the name of the module from ErrorFilesClass to ErrorFileClass
' 10/6/19
'   Generalized the references to errors to messages
'   Changed the class name to MessageFileClass
' 10/11/19
'   Added ReportError for exceptions
' 12/30/19
'   Added ability to specify foldername
'   Omit the closing message for Class and Basic module files
'   Optionally overwrite the file if it exists

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
'   Dim MyMessageFile as MessageFileClass
'   Set MyMessageFile = new MessageFileClass
'
' Report an error by calling the WriteErrorLine routine:
'   MyErrorFile.WriteMessageLine "This is an error message"
'
' Error messages can span multiple lines of text:
'   MyErrorFile.WriteMessgeLine "Line 1" & vbCrLf _
'       "Line 2" & vbCrLf _
'       "Line 3"
'
' Write a blank line to the error file with:
'   MyErrorFile.WriteMessageErrorLines
' To write multiple blank lines, provide a number.
'
' When MyMessageFile goes out of scope, the Class_Terminate
' routine fires and reports the existence and location of the error
' file, if an error was reported.
' If no error was reported, there is no error file and no user message.

Private Const Module_Name As String = "MessageFileClass."

Private pMessageStream As TextStream
Private pStreamName As String
Private pMessageFolderPath As String

Public Property Get MessageFolderPath() As String
    MessageFolderPath = pMessageFolderPath
End Property

Public Sub WriteMessageLine( _
       ByVal Message As String, _
       ByVal StreamName As String, _
       Optional ByVal FolderName As String = vbNullString, _
       Optional ByVal OverWrite As Boolean = False)

    ' This routine writes one line to pErrorStream
    
    Const RoutineName As String = Module_Name & "WriteMessageLine"
    On Error GoTo ErrorHandler
    
    FolderName = IIf(FolderName = vbNullString, StreamName, FolderName)
    
    If pMessageStream Is Nothing Then
        CreateMessageStream StreamName, FolderName, OverWrite
        pStreamName = StreamName
    End If
    
    pMessageStream.WriteLine Message
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' WriteMessageLine

Public Sub WriteBlankMessageLines( _
       ByVal StreamName As String, _
       Optional ByVal NumLines As Long = 1)

    ' This routine writes NumLines blank lines to pErrorStream
    
    Const RoutineName As String = Module_Name & "WriteBlankMessageLines"
    On Error GoTo ErrorHandler
    
    If pMessageStream Is Nothing Then
        CreateMessageStream StreamName
        pStreamName = StreamName
    End If
    
    pMessageStream.WriteBlankLines NumLines
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' WriteBlankMessageLines

Private Sub CreateMessageStream( _
    ByVal StreamName As String, _
    Optional ByVal FolderName As String, _
    Optional ByVal OverWrite As Boolean = False)

    ' This routine creates an Error Message folder
    ' on the user's desktop if one does not already exist.
    ' It then finds the next available Error Message File number and opens it
    
    Const RoutineName As String = Module_Name & "CreateMessageStream"
    On Error GoTo ErrorHandler
    
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    Dim DF As String
    DF = DesktopFolder
    
    pMessageFolderPath = DF & Application.PathSeparator & FolderName
    
    If FSO.FolderExists(pMessageFolderPath) Then
        ' Folder exists. Do nothing.
    Else
        ' Folder does not exist. Create it.
        FSO.CreateFolder pMessageFolderPath
    End If
   ' Message folder now exists on Desktop
   
    Dim FileName As String
    Dim FileNumber As Long
    FileNumber = 1
    
    If OverWrite Then
       FileName = pMessageFolderPath & Application.PathSeparator & StreamName
       
       If FSO.FileExists(FileName) Then
           FSO.DeleteFile FileName, True
       End If
    Else
       While FSO.FileExists(pMessageFolderPath & Application.PathSeparator & StreamName & FileNumber & ".txt")
           FileNumber = FileNumber + 1
       Wend
    
       FileName = pMessageFolderPath & Application.PathSeparator & StreamName & FileNumber & ".txt"
    End If
    
    Set pMessageStream = FSO.CreateTextFile(FileName)
 
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' CreateMessageStream

Private Sub Class_Terminate()
    ' Tell the user if a message file was created
    If Right(pStreamName, 4) = ".cls" Or Right(pStreamName, 4) = ".bas" Then
        pMessageStream.Close
        Set pMessageStream = Nothing
        GoTo Done
    End If
    
    If Not pMessageStream Is Nothing Then
        MsgBox "Check the " & pStreamName & " folder on your Desktop." & vbCrLf _
             & "See the " & pStreamName & " file in that folder " & _
               "for a list of the " & pStreamName & " found.", _
               vbOKOnly Or vbExclamation, _
               pStreamName & " Messages Found"
        pMessageStream.Close
        Set pMessageStream = Nothing
    End If
Done:
End Sub ' Class_Terminate


