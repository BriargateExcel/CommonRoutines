Attribute VB_Name = "CodeImportExportMenus"
'
'
'
' From: https://github.com/spences10/VBA-IDE-Code-Export
'
'
'
Option Explicit

Private MnuEvt      As clsVBECmdHandler
Private EvtHandlers As New Collection

Public Sub auto_open()
    '    CreateVBEMenu
    CreateXLMenu
End Sub                                          ' auto_open

Public Sub auto_close()
    RemoveVBEAndXLMenus
End Sub                                          ' auto_close

Private Sub CreateVBEMenu()
    
    Dim objMenu As CommandBarPopup
    Set objMenu = Application.VBE.CommandBars(1).controls.Add(Type:=msoControlPopup)
    
    With objMenu
        .Caption = "E&xport for VCS"

        Dim objMenuItem As Object
        Set objMenuItem = .controls.Add(Type:=msoControlButton)
        
        objMenuItem.OnAction = "MakeConfigFile"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Make Config File"

        Set objMenuItem = .controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "Export"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Export"

        Set objMenuItem = .controls.Add(Type:=msoControlButton)
        objMenuItem.OnAction = "Import"
        MenuEvents objMenuItem
        objMenuItem.Caption = "&Import"

    End With

    Set objMenuItem = Nothing
    Set objMenu = Nothing

End Sub                                          ' CreateVBEMenu

Private Sub MenuEvents(ByVal objMenuItem As Object)

    Set MnuEvt = New clsVBECmdHandler
    Set MnuEvt.EvtHandler = Application.VBE.Events.CommandBarEvents(objMenuItem)
    EvtHandlers.Add MnuEvt

End Sub                                          ' MenuEvents

Private Sub CreateXLMenu()
    MenuBars(xlWorksheet).Menus.Add Caption:="E&xport for VCS"
    With MenuBars(xlWorksheet).Menus("Export for VCS").MenuItems
        .Add Caption:="&Make Config File", OnAction:="MakeConfigFile"
        .Add Caption:="&Export", OnAction:="Export"
        .Add Caption:="&Import", OnAction:="Import"
    End With

End Sub                                          ' CreateXLMenu

Private Sub RemoveVBEAndXLMenus()

    On Error Resume Next

    Application.VBE.CommandBars(1).controls("Export for VCS").Delete

    '// Clear the EvtHandlers collection if there is anything in it
    While EvtHandlers.Count > 0
        EvtHandlers.Remove 1
    Wend

    Set EvtHandlers = Nothing
    Set MnuEvt = Nothing

    Application.CommandBars("Worksheet Menu Bar").controls("E&xport for VCS").Delete
    On Error GoTo 0

End Sub                                          ' RemoveVBEAndXLMenus

Public Sub btnMakeConfig_onAction(control As IRibbonControl)
    MakeConfigFile
End Sub                                          ' btnMakeConfig_onAction

Public Sub btnExport_onAction(control As IRibbonControl)
    Export
End Sub                                          ' btnExport_onAction

Public Sub btnImport_onAction(control As IRibbonControl)
    Import
End Sub                                          ' btnImport_onAction

