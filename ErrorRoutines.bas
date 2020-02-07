Attribute VB_Name = "ErrorRoutines"
Option Explicit

Private Const Module_Name As String = "ErrorRoutines."
'
' This module provides the error handling routines
' See the example usage at the end of the module
'
Private pErrorFile As ErrorFileClass
Private SourceOfError As String

Public Property Get ErrorsFound() As Boolean
    ErrorsFound = Not pErrorFile Is Nothing
End Property

Public Sub RaiseError( _
       ByVal ErrorNo As Long, _
       ByVal Src As String, _
       ByVal Proc As String, _
       ByVal Desc As String, _
       ParamArray Args() As Variant)

    ' https://excelmacromastery.com/vba-error-handling/
    ' Reraises an error and adds line number and current procedure name
    ' Adds a list of parameter names and corresponding parameter values
    ' One name and value per line

    ' Check if procedure where error occurs the line no and proc details
    ' Add error line number if present
    If Erl <> 0 Then
        SourceOfError = vbCrLf & "Line no: " & Erl & " "
    End If

    ' Add procedure to source
    SourceOfError = SourceOfError & vbCrLf & Proc
    
    Dim I As Long
    For I = 1 To IIf(UBound(Args, 1) Mod 2 = 2, UBound(Args, 1), UBound(Args, 1) - 1) Step 2
        SourceOfError = SourceOfError & Args(I) & " = " & Args(I + 1) & vbCrLf
    Next I

    ' If the code stops here,
    ' make sure DisplayError is placed in the top most Sub
    Err.Raise ErrorNo, SourceOfError, Desc

End Sub                                          ' RaiseError

Public Sub DisplayError(ByVal ProcName As String)

    ' https://excelmacromastery.com/vba-error-handling/
    ' Displays the error when it reaches the topmost sub

    Dim Msg As String
    Msg = "The following error occurred: " & vbCrLf & _
          Err.Description & vbCrLf & _
          Err.Source & vbCrLf & _
          ProcName

    ReportError Msg
    
    CloseErrorFile

End Sub                                          ' DisplayError

Public Sub ReportError( _
       ByVal ErrMsg As String, _
       ParamArray Args() As Variant)

    ' This routine writes an error message to the error file
    
    Const RoutineName As String = Module_Name & "ReportError"
    On Error GoTo ErrorHandler
    
    #If DebugOn Then
        Dim cPerf As PerformanceClass
        If gbDebug(RoutineName) Then
            Set cPerf = New PerformanceClass
            cPerf.SetRoutine RoutineName
        End If
    #End If
    
    If pErrorFile Is Nothing Then
        Set pErrorFile = New ErrorFileClass
    End If
    
    Dim ErrorMessage As String
    Dim I As Long
    For I = 0 To IIf(UBound(Args, 1) Mod 2 = 0, UBound(Args, 1) - 2, UBound(Args, 1) - 1) Step 2
        ErrorMessage = ErrorMessage & Args(I) & " = " & Args(I + 1) & vbCrLf
    Next I

    pErrorFile.WriteErrorLine ErrMsg & vbCrLf & ErrorMessage & vbCrLf
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ReportError

Public Function CloseErrorFile() As Boolean      ' Declared as a function to keep it of the Alt-F8 list of executable routines

    Set pErrorFile = Nothing
    SourceOfError = vbNullString
    
End Function                                     ' CloseErrorFile

'Private Sub SubErrorRaiseProcess(ByVal Parameter As String)
'
'    ' This routine tests the error raise process
'
'    Const RoutineName As String = Module_Name & "SubErrorRaiseProcess"
'    On Error GoTo ErrorHandler
'
'    ReportError "Error Message", _
'        "Parameter", Parameter, _
'        "Param 1", 1, _
'        "Param 2", 2
'
'    Dim Test As Long
'    Test = 1 / 0
'
'    '@Ignore LineLabelNotUsed
'Done:
'    Exit Sub
'ErrorHandler:
'    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
'End Sub                                          ' SubErrorRaiseProcess
'
'Public Sub TestErrorRaiseProcess()
'
'    ' This routine tests the error raise process
'
'    Const RoutineName As String = Module_Name & "TestErrorRaiseProcess"
'    On Error GoTo ErrorHandler
'
'    SubErrorRaiseProcess "First"
'    SubErrorRaiseProcess "Second"
'
'    CloseErrorFile
'
'    '@Ignore LineLabelNotUsed
'Done:
'    Exit Sub
'ErrorHandler:
'    DisplayError RoutineName ' Only in the Main routine; RaiseError in all others
'End Sub                                          ' TestErrorRaiseProcess





