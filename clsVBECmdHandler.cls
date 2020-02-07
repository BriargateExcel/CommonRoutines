VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBECmdHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'
'
' From: https://github.com/spences10/VBA-IDE-Code-Export
'
'
'
Option Explicit

'// Add references for :
'//     Microsoft Visual Basic For Applications Extensibility 5.3

Public WithEvents EvtHandler As VBIDE.CommandBarEvents
Attribute EvtHandler.VB_VarHelpID = -1

Private Sub EvtHandler_Click( _
        ByVal CommandBarControl As Object, _
        ByRef Handled As Boolean, _
        ByRef CancelDefault As Boolean)

    On Error Resume Next

    '// Run the code specified in the object's OnAction property.
    Application.Run CommandBarControl.OnAction

    '// Indicate to the Events object that we've successfully handled the event.
    Handled = True
    CancelDefault = True

    On Error GoTo 0
End Sub


