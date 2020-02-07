Attribute VB_Name = "Templates"
Option Explicit

Private Const Module_Name As String = "Templates."

Public Sub MainProgram()

    ' Used as the top level routine
    
    Const RoutineName As String = Module_Name & "MainProgram"
    On Error GoTo ErrorHandler
    
    ' Code goes here
    
Done:
    MsgBox "Normal exit", vbOKOnly
Halted:
    ' Use the Halted exit point after giving the user a message
    '   describing why processing did not run to completion
    CloseErrorFile
    TurnOnAutomaticProcessing
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    TurnOnAutomaticProcessing
    CloseErrorFile
End Sub ' MainProgram

Private Sub SubTemplate()

    ' Used for lower level routines
    
    Const RoutineName As String = Module_Name & "SubTemplate"
    On Error GoTo ErrorHandler
    
    ' Code goes here
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' SubTemplate

Private Function CheckTemplate( _
    ByVal Parm As String _
    ) As Boolean
    
    ' Used to return a boolean and no other value

    Const RoutineName As String = Module_Name & "CheckTemplate"
    On Error GoTo ErrorHandler
    
    ' Code goes here
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CheckTemplate

Private Function TryTemplate( _
    ByVal Parm1 As String, _
    ByRef Parm2 As String _
    ) As Boolean
    
    ' Used to return a boolean and some other value(s)
    ' Returns True if successful

    Const RoutineName As String = Module_Name & "TryTemplate"
    On Error GoTo ErrorHandler
    
    ' Code goes here
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryTemplate



