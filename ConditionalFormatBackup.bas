Attribute VB_Name = "ConditionalFormatBackup"
Option Explicit

Public Sub CFBackup()
    '
    '   Macro to backup the ACTIVE SHEET's conditional formatting (CF) using named ranges that auto-adjust to row/column changes
    '   + Create a new backup sheet
    '   + Define a named range with Workbook scope relating the backup sheet to ActiveSheet
    '       - A named range will auto-adjust if ActiveSheet is renamed
    '   + Add a defined name with ActiveSheet scope to register information about the backup
    '   + Define a named range with ActiveSheet scope for each CF rule's AppliesTo range
    '       - A named range will auto-adjust if rows or columns are inserted or removed
    '   + For each CF rule:
    '       - Locate the first cell in the rule's AppliesTo range
    '       - If the first cell is merged, copy it to a temporary cell before it is unmerged
    '       - Copy the first cell's CF rule(s) to the backup sheet
    '       - Isolate one unique rule from the CF rule(s) copied to the backup sheet
    '       - Record the rule's defined name, range formulas, etc., on the backup sheet
    '   + Protect the backup sheet
    '
    '   Feb 2019 by J. Woolley
    '

    Const myName As String = "CFBackup"

    Const sPre As String = "\" + myName + "\"
    Const sPreItem As String = sPre + "Item\"
    Const sPreSheet As String = sPre + "Sheet\"

    Const sFormat As String = "0000"
    Const sSignet As String = "HIDE THIS ROW"    ' used as backup sheet signature
    Const sLabel As String = vbNullString        ' see Function ProgressBar_Text
    Const sDND As String = "DO NOT DISTURB "

    Const bVisible As Boolean = False
    Const nRowZero As Long = 4

    ' Create a new backup sheet
    Dim oSheet As Worksheet
    Set oSheet = ActiveWorkbook.ActiveSheet

    Dim sName As String
    sName = sPreSheet + sFormat

    Dim oName As Name
    Set oName = oSheet.Names.Add(sName, "=$A$1", bVisible) ' temporary Name to get qualified scope

    sName = oSheet.Name + "!" + sName            ' unqualified scope plus Name

    Dim sScope As String
    sScope = NameRefersTo(sName, -1) + "!"       ' qualified scope

    oName.Delete

    Dim nItems As Long
    nItems = oSheet.Cells.FormatConditions.Count

    If nItems < 1 Then

        Dim Msg As String
        Msg = "The ACTIVE SHEET has NO conditional formatting rules to back up."

        MsgBox Msg, (vbOKOnly + vbExclamation), myName
        Exit Sub
    End If

    Dim sWord As String
    sWord = IIf(nItems = 1, "rule", "rules")

    Msg = "This macro will back up the ACTIVE SHEET's conditional formatting (CF) " + _
          sWord + ", but first make sure its current CF " + sWord + _
          IIf(nItems = 1, " is", " are") + _
          " correct in the Conditional Formatting Rules Manager (Alt+H+L+R). " + _
          vbNewLine + _
          "Progress will be reported in Excel's status bar at bottom-left." + _
          vbNewLine + vbNewLine + _
          "Do you want to back up the ACTIVE SHEET's " + CStr(nItems) + _
          " conditional formatting " + sWord + " now?"

    Dim sBackup As String
    sBackup = vbNullString

    sName = sScope + sPreItem + sFormat
    Set oName = Nothing
    On Error Resume Next
    Set oName = oSheet.Names(sName)              ' check for previous backup
    On Error GoTo 0

    If Not (oName Is Nothing) Then

        Dim vInfo As Variant
        vInfo = Split(oName.RefersTo)

        vInfo(0) = Mid$(vInfo(0), 3)

        Dim n As Long
        n = CInt(Mid$(vInfo(3), Len("Count=X"), Len(sFormat))) ' literal "Count=" must match the restore macro

        Msg = _
            Msg + " If so, the previous backup of " + CStr(n) + _
            IIf(n = 1, " rule", " rules") + _
            " on " + vInfo(0) + " at " + vInfo(1) + " " + vInfo(2) + " will be replaced."

        Dim sSelec As String
        sSelec = Mid$(vInfo(4), Len("Sheet=X"), Len(sFormat)) ' literal "Sheet=" must match the restore macro

        Dim sNamSheet As String
        sNamSheet = sPreSheet + sSelec

        If (CStr(NameRefersTo(sNamSheet, -1)) + "!") = sScope Then sBackup = myName + sSelec

    End If

    Dim ans As Variant
    ans = MsgBox(Msg, (vbYesNo + vbQuestion + vbDefaultButton2), myName)

    If ans <> vbYes Then Exit Sub

    If sBackup = vbNullString Then               ' find next available backup sheet name
        n = 0

        Do While True
            n = n + 1
            sNamSheet = sPreSheet + Format$(n, sFormat)

            If IsError(NameRefersTo(sNamSheet, -1)) Then ' available
                sBackup = myName + Format$(n, sFormat)
                Exit Do
            End If

        Loop

        Dim oBackup As Worksheet
        Set oBackup = Nothing

        On Error Resume Next
        Set oBackup = ActiveWorkbook.Worksheets(sBackup)
        On Error GoTo 0

        If Not (oBackup Is Nothing) Then

            If oBackup.Cells((nRowZero - 1), 1).Value <> sSignet Then
                Msg = "The '" + sBackup + "' sheet does not have the expected signature but will be replaced. " _
                    + "Quit and rename it if necessary." + vbNewLine + vbNewLine _
                    + "Click OK to continue or Cancel to quit."

                ans = MsgBox(Msg, (vbOKCancel + vbInformation + vbDefaultButton2), myName)
                If ans <> vbOK Then Exit Sub
            End If

        End If

        Set oName = ActiveWorkbook.Names.Add(sNamSheet, ("=" + sScope + "$A$1"), bVisible)
        oName.Comment = sDND + sPreSheet + "... names; see macro " + myName
    End If

    With Application

        Dim bDisplayStatusBar As Boolean
        bDisplayStatusBar = .DisplayStatusBar    ' save original show/hide

        .DisplayStatusBar = True
        .ScreenUpdating = False
    End With

    sSelec = Selection.Address                   ' save original selection address

    Dim nLastRow As Long
    nLastRow = oSheet.UsedRange.Row - 1 + oSheet.UsedRange.Rows.Count ' last used row

    Dim nSteps As Long
    nSteps = oSheet.Names.Count + nItems * 2

    Dim nStep As Long
    nStep = 0

    sName = sScope + sPre + "*"

    For Each oName In oSheet.Names
        nStep = nStep + 1
        Application.StatusBar = ProgressBar_Text(sLabel, nStep, nSteps)
        If oName.Name Like sName Then oName.Delete ' delete any previous backup defined names
    Next oName

    sName = sScope + sPreItem + sFormat          ' create new defined names with backup data

    Dim sNow As String
    sNow = Now

    vInfo = Split(sNow)

    ' literals "Count=" and "Sheet=" must match the restore macro
    Set oName = _
              oSheet.Names.Add(sName, (sNow + " Count=" + Format$(nItems, sFormat) + _
                                       " Sheet=" + Right$(sBackup, Len(sFormat))), bVisible)

    oName.Comment = sDND + sPreItem + "... names; see macro " + myName

    For n = 1 To nItems
        nStep = nStep + 1
        Application.StatusBar = ProgressBar_Text(sLabel, nStep, nSteps)
        sName = sScope + sPreItem + Format$(n, sFormat)

        Dim oFC As Object
        Set oFC = oSheet.Cells.FormatConditions(n)

        Set oName = oSheet.Names.Add(sName, ("=" + sScope + oFC.AppliesTo.Address), bVisible)
        oName.Comment = sDND
    Next n

    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets(sBackup).Delete    ' delete any previous sheet with the same name
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set oBackup = ActiveWorkbook.Worksheets.Add(after:=oSheet)
    ' Backup sheet created

    With oBackup                                 ' record backup results
        .Name = sBackup
        .Tab.Color = RGB(255, 0, 0)
        n = nRowZero - 3                         ' top row

        With .Rows(n)
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(255, 0, 0)
        End With

        With .Range(.Rows(n), .Rows(n + 1))
            .RowHeight = .RowHeight * 1.5
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With

        n = nRowZero - 1                         ' hidden row
        .Rows(n).Hidden = True

        Dim sCell2 As String
        sCell2 = .Cells(n, 2).Address

        Dim sCell3 As String
        sCell3 = .Cells(n, 3).Address

        Dim sCell4 As String
        sCell4 = .Cells(n, 4).Address

        '        With .Range(Cells(nRowZero, 1), Cells(nRowZero, 7))
        With .Range(ActiveSheet.Cells(nRowZero, 1), ActiveSheet.Cells(nRowZero, 7))
            .Font.Bold = True

            For n = xlEdgeLeft To xlInsideVertical
                With .Borders(n)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With

            Next n

        End With

        n = nRowZero - 3                         ' top row
        .Cells(n, 1).Formula = _
                             "=IF(ISERROR(" + sCell2 + "),""The " + myName _
                           + " macro created this backup for a worksheet that no longer exists; therefore, this ""&" + _
                             sCell4 + "&"" is invalid."",""" + _
                                                         sDND + "this ""&" + sCell4 + _
                                                         "&"" created by the " + myName + _
                                                         " macro. Deleting the '""&" + sCell3 + _
                                                         "&""' worksheet invalidates this backup."")"

        n = nRowZero - 2                         ' 2nd row
        .Cells(n, 1).Formula = _
                             "=""Backup of " + CStr(nItems) + _
                             " conditional formatting (CF) " + sWord + _
                             " on " + vInfo(0) + " at " + vInfo(1) + _
                             " " + vInfo(2) + _
                             " for ""&IF(ISERROR(" + sCell2 + "),""a worksheet originally named '" + oSheet.Name + "' but subsequently deleted."",""the '""&" + sCell3 + "&""' worksheet."")"


        n = nRowZero - 1                         ' hidden row
        .Cells(n, 1).Value = sSignet             ' backup sheet signature
        .Cells(n, 2).Formula = "=PERSONAL.xlsb!NameRefersTo(""" + sNamSheet + """,-1)&""!"""
        .Cells(n, 3).Formula = "=PERSONAL.xlsb!NameRefersTo(""" + sNamSheet + """,-2)"
        .Cells(n, 4).Formula = "=IF(PERSONAL.xlsb!IsProtected(),""protected backup (no password)"",""unprotected backup"")"
        .Cells(n, 7).Value = "F9 to recalc"

        n = nRowZero                             ' heading row
        .Cells(n, 1).Value = "Defined Name"
        .Cells(n, 2).Value = "Name 'Refers To' Range"
        .Cells(n, 3).Value = "CF 'Applies To' Range"
        .Cells(n, 4).Value = "CF Sample"
        .Cells(n, 5).Value = "CF Type"
        .Cells(n, 6).Value = "CF Type Description"
        .Cells(n, 7).Value = "Stop If True"

        .Columns(7).HorizontalAlignment = xlHAlignCenter
        .Calculate

        Dim vType As Variant
        vType = Array("NA", "Cell Value", "Expression", "Color Scale", _
                      "Databar", "Top 10", "Icon Set", "NA", "Unique Values", _
                      "Text String", "Blanks Condition", "Time Period", _
                      "Above Average Condition" _
                      , "No Blanks Condition", "NA", "NA", "Errors Condition", _
                      "No Errors Condition")

        Dim nRow As Long
        nRow = nRowZero
        ' Backup sheet set up

        ' Populate backup sheet
        For n = 1 To nItems
            nStep = nStep + 1
            Application.StatusBar = ProgressBar_Text(sLabel, nStep, nSteps)
            nRow = nRow + 1
            sName = sCell2 + "&" + .Cells(nRow, 1).Address
            Set oFC = oSheet.Cells.FormatConditions(n)
            .Cells(nRow, 1).Value = sPreItem + Format$(n, sFormat)
            .Cells(nRow, 2).Formula = "=PERSONAL.xlsb!NameRefersTo(" + sName + ",0)"
            .Cells(nRow, 3).Formula = "=HYPERLINK(""#""&Personal.xlsb!NameRefersTo(" + sName + ",2),Personal.xlsb!NameRefersTo(" + sName + ",1))"
            .Cells(nRow, 4).Value = vbNullString
            .Cells(nRow, 5).Value = oFC.Type
            .Cells(nRow, 6).Value = vType(oFC.Type)
            .Cells(nRow, 7).Value = IIf(oFC.StopIfTrue, "X", vbNullString)
            sName = sScope + sPreItem + Format$(n, sFormat)

            Dim rFirst As Range
            Set rFirst = ActiveSheet.Range(sName).Areas(1).Item(1)

            Dim bMerge As Boolean
            bMerge = rFirst.MergeCells

            If bMerge Then                       ' treat merged cells

                Dim rMerge As Range
                Set rMerge = rFirst.MergeArea

                Application.DisplayAlerts = False
                rMerge.Copy                      ' copy to temporary
                rMerge.Offset(nLastRow).PasteSpecial Paste:=xlPasteAll
                Application.DisplayAlerts = True
                Set rFirst = rFirst.Offset(nLastRow)
                Set rMerge = rFirst.MergeArea
                rMerge.UnMerge                   ' isolate first cell
            End If

            Application.DisplayAlerts = False
            rFirst.Copy                          ' copy conditional format(s)
            .Cells(nRow, 4).PasteSpecial Paste:=xlPasteAll
            Application.DisplayAlerts = True

            .Cells(nRow, 8).Copy                 ' copy default format (blank cell)
            .Cells(nRow, 4).PasteSpecial Paste:=xlPasteAllMergingConditionalFormats
            .Cells(nRow, 4).Value = "CF Rule #" + CStr(n)

            If bMerge Then
                rMerge.Merge
                rMerge.Delete xlShiftUp          ' delete temporary
            End If

            Dim nTotal As Long
            nTotal = .Cells(nRow, 4).FormatConditions.Count ' new CF rules are added at beginning (#1 of Count)

            If nTotal > 1 Then                   ' isolate single CF rule for this cell
                Set rFirst = ActiveSheet.Range(sName).Areas(1).Item(1)

                Dim nCount As Long
                nCount = 0

                Dim I As Long
                For I = (nRowZero + 1) To (nRow - 1) ' count previous CF rules intersecting with this CF rule
                    sName = .Range(sCell2).Value + .Cells(I, 1).Value

                    Dim rSect As Range
                    Set rSect = Application.Intersect(ActiveSheet.Range(sName), rFirst)

                    If Not (rSect Is Nothing) Then nCount = nCount + 1
                Next I

                For I = 1 To nCount              ' delete previous CF rules
                    .Cells(nRow, 4).FormatConditions(1).Delete
                    nTotal = nTotal - 1
                Next I

                For I = 2 To nTotal              ' delete any CF rules that follow
                    .Cells(nRow, 4).FormatConditions(2).Delete
                Next I

            End If

        Next n
        ' All rows in the table are now set up

        ' Clean up the appearance of the backup sheet
        .Calculate
        .Activate
        .Range(.Cells(nRowZero, 1), .Cells(nRow, 7)).Columns.AutoFit
        .Cells((nRowZero + 1), 1).Select
        ActiveWindow.FreezePanes = True
        .Visible = xlSheetHidden
        .Protect AllowFormattingColumns:=True    ' no password
        .Calculate
    End With                                     ' oBackup

    oSheet.Activate
    ActiveSheet.Range(sSelec).Select             ' restore original selection

    With Application
        .CutCopyMode = False                     ' cancel any moving border
        .StatusBar = False                       ' restore default status bar
        .DisplayStatusBar = bDisplayStatusBar    ' restore original show/hide
        .ScreenUpdating = True
    End With                                     ' oBackup

    Msg = _
        "A backup of the ACTIVE SHEET's " + CStr(nItems) + _
        " conditional formatting (CF) " + sWord + " was created on " + _
        vInfo(0) + " at " + vInfo(1) + " " + vInfo(2) + _
        " using defined names '" + sPreItem + "...' with worksheet " + _
        "scope, so each 'Applies To' range will automatically adjust if rows or columns are inserted or removed. " + _
        IIf(bVisible, vbNullString, "These defined names are hidden from Name Manager (Ctrl+F3) and Go To (F5). ") + _
        vbNewLine + vbNewLine + _
        "A new hidden '" + oBackup.Name + "' sheet was created to record results. " + _
        "Do you want to make that sheet visible?"

    ans = MsgBox(Msg, (vbYesNo + vbQuestion + vbDefaultButton2), myName)

    If ans = vbYes Then
        oBackup.Visible = xlSheetVisible
        oBackup.Activate
    End If

End Sub

Public Sub CFRestore()
    '
    '   Macro to restore the ACTIVE SHEET's conditional formatting (CF) using a backup created by the CFBackup macro
    '   + Delete all CF rules on ActiveSheet
    '   + For each unique CF rule on the backup sheet:
    '       - Copy the CF rule to a temporary cell on ActiveSheet
    '       - Modify the temporary cell's AppliesTo range to match the CF rule's named range (auto-adjusted)
    '       - Delete the temporary cell
    '   + Protect the backup sheet (if appropriate)
    '
    '   Feb 2019 by J. Woolley
    '

    Const myName As String = "CFRestore"
    Const sBacMac As String = "CFBackup"         ' the following Constants must match the backup macro
    Const sPreItem As String = "\" + sBacMac + "\Item\"
    Const sPreSheet As String = "\" + sBacMac + "\Sheet\"
    Const sFormat As String = "0000"
    Const sLabel As String = vbNullString
    Const nRowZero As Long = 4

    Dim oSheet As Worksheet
    Set oSheet = ActiveWorkbook.ActiveSheet

    Dim sName As String
    sName = sPreSheet + sFormat

    Dim oName As Name
    Set oName = oSheet.Names.Add(sName, "=$A$1", False) ' temporary Name to get qualified scope

    sName = oSheet.Name + "!" + sName            ' unqualified scope plus Name

    Dim sScope As String
    sScope = NameRefersTo(sName, -1) + "!"       ' qualified scope

    oName.Delete

    Dim sWord As String
    sWord = "; therefore, the ACTIVE SHEET's conditional formatting rules cannot be restored."

    sName = sScope + sPreItem + sFormat
    Set oName = Nothing

    On Error Resume Next
    Set oName = oSheet.Names(sName)              ' check for previous backup
    On Error GoTo 0

    If (oName Is Nothing) Then

        Dim Msg As String
        Msg = sBacMac + " data is not registered for the ACTIVE SHEET." + vbNewLine + vbNewLine _
            + sName + " could not be found in the collection of defined names" + sWord

        MsgBox Msg, vbExclamation, myName
        Exit Sub
    End If

    Dim vInfo As Variant
    vInfo = Split(oName.RefersTo)

    vInfo(0) = Mid$(vInfo(0), 3)

    Dim nItems As Long
    nItems = CInt(Mid$(vInfo(3), Len("Count=X"), Len(sFormat))) ' literal "Count=" must match the backup macro

    Dim sSelec As String
    sSelec = Mid$(vInfo(4), Len("Sheet=X"), Len(sFormat)) ' literal "Sheet=" must match the backup macro

    Dim sBackup As String
    sBackup = sBacMac + sSelec

    sName = sPreSheet + sSelec

    Dim ans As Variant
    ans = NameRefersTo(sName, -1)

    sSelec = CStr(ans)

    If (sSelec + "!") <> sScope Then
        Msg = _
            "The ACTIVE SHEET's backup was previously assigned to the '" + _
            sBackup + "' sheet but is no longer applicable"

        If Not IsError(ans) Then
            Msg = Msg + ". That backup applies to the '" + _
                  NameRefersTo(sName, -2) + "' sheet instead"
        End If

        Msg = Msg + sWord
        MsgBox Msg, vbExclamation, myName
        Exit Sub
    End If

    Dim oBackup As Worksheet
    Set oBackup = Nothing

    On Error Resume Next
    Set oBackup = ActiveWorkbook.Worksheets(sBackup)
    On Error GoTo 0

    If (oBackup Is Nothing) Then
        Msg = "The '" + sBackup + "' sheet could not be found" + sWord

        MsgBox Msg, vbExclamation, myName
        Exit Sub
    End If

    sWord = IIf(nItems = 1, "rule", "rules")

    Msg = "A backup of the ACTIVE SHEET's conditional formatting (CF) " + _
          sWord + " was created on " + vInfo(0) + " at " + vInfo(1) + _
          " " + vInfo(2) + "." + vbNewLine + vbNewLine + _
          "Do you want to restore that backup's " + CStr(nItems) + _
          " " + sWord + " now?"

    Dim n As Long
    n = oSheet.Cells.FormatConditions.Count

    If n > 0 Then
        Msg = Msg + " If so, the " + CStr(n) + _
                                             " current CF " + IIf(n = 1, "rule", "rules") + _
                                             " will be replaced."
    End If

    Msg = Msg + " Progress will be reported in Excel's status bar at bottom-left."

    ans = MsgBox(Msg, (vbYesNo + vbQuestion), myName)
    If ans <> vbYes Then Exit Sub

    With Application

        Dim bDisplayStatusBar As Boolean
        bDisplayStatusBar = .DisplayStatusBar    ' save original show/hide

        .DisplayStatusBar = True
        .ScreenUpdating = False
    End With

    sSelec = Selection.Address                   ' save original selection address

    Dim nNextRow As Long
    nNextRow = oSheet.UsedRange.Row + oSheet.UsedRange.Rows.Count ' first unused row

    Dim oFCs As FormatConditions
    Set oFCs = oSheet.Cells.FormatConditions

    oFCs.Delete

    With oBackup
        .Calculate

        Dim bProtect As Boolean
        bProtect = .ProtectContents              ' save original protection

        If bProtect Then .Unprotect

        Dim nRow As Long
        nRow = nRowZero + nItems                 ' last-to-first

        With .Range(.Cells(nRowZero, 1), .Cells(nRow, 7))
            .Sort Key1:=.Cells(nRowZero, 1), Order1:=xlAscending, Header:=xlYes
            .Columns.AutoFit
        End With

        Dim nSkip As Long
        nSkip = 0

        For n = 1 To nItems
            Application.StatusBar = ProgressBar_Text(sLabel, n, nItems)

            If IsError(.Cells(nRow, 3)) Then
                nSkip = nSkip + 1
            Else
                sName = .Cells(nRow, 3).Value
                .Cells(nRow, 4).Copy             ' add this CF at oFCs(1) shifting previous to 2, 3, etc.

                With oSheet.Cells(nNextRow, 1)   ' temporary cell
                    .PasteSpecial Paste:=xlPasteAll
                    oFCs(1).ModifyAppliesToRange Range:=ActiveSheet.Range(sName)
                    .Delete xlShiftUp
                End With

            End If

            nRow = nRow - 1

        Next n

        If bProtect Then
            .Protect AllowFormattingColumns:=True ' restore original protection
            .Calculate
        End If

    End With

    ActiveSheet.Range(sSelec).Select             ' restore original selection

    With Application
        .CutCopyMode = False                     ' cancel any moving border
        .StatusBar = False                       ' restore default status bar
        .DisplayStatusBar = bDisplayStatusBar    ' restore original show/hide
        .ScreenUpdating = True
    End With

    Msg = _
        CStr(nItems - nSkip) + " of " + CStr(nItems) + _
        " conditional formatting " + sWord + IIf((nItems - nSkip) = 1, " was", " were") + _
        " restored to the ACTIVE SHEET from the backup created on " + _
        vInfo(0) + " at " + vInfo(1) + " " + vInfo(2) + "."

    If nSkip > 0 Then
        Msg = Msg + " (" + CStr(nSkip) + IIf(nSkip = 1, " rule was", " rules were") + " previously deleted.)"
    End If

    Msg = _
        Msg + vbNewLine + vbNewLine + _
        "Results can be reviewed using the Conditional Formatting Rules Manager (Alt+H+L+R)."

    MsgBox Msg, (vbOKOnly + vbInformation), myName

End Sub

Public Sub CFBackupAll()
    '
    '   This macro will offer to backup conditional formatting for ALL WORKSHEETS of the ACTIVE WORKBOOK.
    '
    '   Mar 2019 by J. Woolley
    '

    Const myName As String = "CFBackupAll"
    Const sBacMac As String = "CFBackup"         ' the following Constants must match the backup macro
    Const sSignet As String = "HIDE THIS ROW"    ' used as backup sheet signature
    Const nRowZero As Long = 4

    Dim Msg As String
    Msg = "Do you want to consider backup of conditional formatting (CF) for ALL WORKSHEETS of the ACTIVE WORKBOOK?"
    Dim ans As Variant

    ans = MsgBox(Msg, (vbQuestion + vbYesNo), myName)
    If ans <> vbYes Then Exit Sub

    Dim oRestoreWS As Worksheet
    Set oRestoreWS = ActiveSheet

    Dim nShts As Long                            ' number of sheets
    nShts = 0                                    ' number of sheets

    Dim nDone As Long                            ' number done
    nDone = 0                                    ' number of sheets updated

    Msg = vbNullString
    Dim oWS As Worksheet
    For Each oWS In ActiveWorkbook.Worksheets    ' process each worksheet
        If Not (oWS.Name Like (sBacMac + "####")) Or oWS.Cells((nRowZero - 1), 1) <> sSignet Then
            nShts = nShts + 1

            Dim nFCs As Long                     ' number of sheet's FormatConditions
            nFCs = oWS.Cells.FormatConditions.Count

            Msg = Msg + "Worksheet '" + oWS.Name + "' has " + CStr(nFCs) + " CF rules"

            If nFCs < 1 Then
                Msg = Msg + "." + vbNewLine + vbNewLine
            ElseIf oWS.Visible <> xlSheetVisible Then
                Msg = Msg + " but is hidden." + vbNewLine + vbNewLine
            ElseIf oWS.ProtectContents Then
                Msg = Msg + " but is protected." + vbNewLine + vbNewLine
            Else
                oWS.Activate
                Msg = Msg + "."
                MsgBox Msg, (vbInformation + vbOKOnly), myName
                Msg = vbNullString
                CFBackup
                nDone = nDone + 1
            End If

        End If

    Next oWS

    oRestoreWS.Activate

    Msg = _
        Msg + "Conditional formatting backup was considered for " + _
        CStr(nDone) + " of " + CStr(nShts) + _
        " worksheets in the ACTIVE WORKBOOK."

    MsgBox Msg, vbInformation, myName

End Sub

Public Function NameRefersTo( _
       ByVal Name As String, _
       Optional ByVal Choice As Long = 0 _
       ) As Variant
    '
    '   User-Defined Function (UDF) to return the RefersTo property (as String) of the defined name (named range) Name
    '       The initial equals-sign (=) will be removed from the returned String
    '       If Name has Worksheet scope, then Name must reference that scope; e.g., MySheet!MyName or 'My Sheet'!My_Name
    '       If Name has Workbook scope, then Name must NOT reference a Worksheet; e.g., MyName or My_Name
    '       When this function is used in a cell's formula, Name is assumed to be defined in that cell's Workbook
    '       When this function is used in a VBA statement, Name is assumed to be defined in ActiveWorkbook
    '       If Name is defined in another Workbook, then Name must reference that Workbook (which must be open);
    '           e.g.,[MyBook.xlsx]MyName or [My Book.xlsx]My_Name or [My Book.xlsx]!My_Name
    '           or [My Book.xlsx]MySheet!MyName or [My Book.xlsx]'My Sheet'!My_Name or '[My Book.xlsx]My Sheet'!My_Name
    '       If Name is not a valid defined name, #VALUE! (Error 2015) will be returned as Variant
    '   If Choice =  0 (default), the RefersTo property will be returned directly
    '   If Choice = -1 and RefersTo represents a Range, only the qualified Worksheet reference (see below) will be returned
    '       In this case, #VALUE! (Error 2015) will be returned if the reference does not represent a valid Worksheet
    '       Single-quotes necessary to qualify a Worksheet reference will be retained in the returned string; e.g., 'JW''s Sheet'
    '       The Worksheet reference might include a Workbook in square-brackets; e.g., [MyBook.xlsx]MySheet or '[My Book.xlsx]My Sheet'
    '       If the named range includes more than one Area, only the first Area's Worksheet reference will be returned (see below)
    '       A trailing exclamation-point (!) will NOT be included in the returned Worksheet reference
    '       A null string ("") will be returned if RefersTo does not include a Worksheet reference (unexpected for a named range)
    '   If Choice < -1 and RefersTo represents a Range, only the reference's unqualified Worksheet name will be returned
    '       In this case, #VALUE! (Error 2015) will be returned if the reference does not represent a valid Worksheet
    '       Single-quotes necessary to qualify a Worksheet name will NOT be retained in the returned string; e.g., JW's Sheet
    '       The reference's Workbook (if specified) will NOT be included in the returned string
    '       If the named range includes more than one Area, only the first Area's Worksheet name will be returned (see below)
    '       A trailing exclamation-point (!) will NOT be included in the returned Worksheet name
    '       ActiveSheet.Name will be returned if RefersTo does not include a Worksheet reference (unexpected for a named range)
    '   If Choice =  1 and RefersTo represents a Range, all Worksheet references PLUS any #REF! will be removed from the returned String
    '       In this case, #VALUE! (Error 2015) will be returned if RefersTo minus any #REF! does not represent a valid Range
    '       If the named range includes more than one Area, all Areas must have the same Worksheet reference (see below)
    '   If Choice >  1 and RefersTo represents a Range, only #REF! and RELATED Worksheet references will be removed from the returned String
    '       In this case, #VALUE! (Error 2015) will be returned if RefersTo minus any #REF! does not represent a valid Range
    '       If the named range includes more than one Area, all Areas must have the same Worksheet reference (see below)
    '   If Choice <> 0 and RefersTo references a Workbook that is not open, #VALUE! (Error 2015) will be returned
    '
    '   The Worksheet reference in Workbook.Names(Name).RefersTo might not be the same as the scope of Workbook.Names(Name)
    '   Isolating parts of a Worksheet reference is difficult because a Worksheet name can include almost any character (not [ or ])
    '   It is unusual (but possible) for a named range with more than one Area to have different Worksheet references
    '   But a named range with multiple Areas having different Worksheet references will not represent a valid Range
    '   So if Choice <> 0, a named range with multiple Areas having different Worksheet references will return #VALUE! (Error 2015)
    '
    '   Feb 2019 by J. Woolley
    '

    Application.Volatile                         ' RefersTo is not static for a given Name

    Dim oWB As Workbook
    Set oWB = ActiveWorkbook                     ' when this function is referenced in a VBA statement
    On Error Resume Next                         ' in case ThisCell is not valid
    Set oWB = Application.ThisCell.Parent.Parent ' when this function is referenced in a cell's formula
    On Error GoTo Return_Error                   ' in case Name is not valid or RefersTo does not represent a Range

    Dim sBook As String
    sBook = Mid$(Name, (InStr(1, Name, "[") + 1)) ' extract yyy]zzz from x[yyy]zzz (if present)

    Dim TempName As String
    TempName = Name

    If sBook <> Name Then
        sBook = Left$(sBook, InStr(1, sBook, "]")) ' extract yyy] from yyy]zzz

        If sBook <> vbNullString Then
            sBook = Left$(sBook, (Len(sBook) - 1)) ' extract yyy from yyy]
            Set oWB = Workbooks(sBook)           ' Workbook related to Name (error if not open)

            Dim sResult As String
            sResult = Mid$(Name, (InStr(1, Name, "]") + 1)) ' extract zzz from x[yyy]zzz

            If Left$(sResult, 1) = "!" Or Left$(sResult, 2) = "'!" Then
                sResult = Mid$(sResult, (InStr(1, sResult, "!") + 1))
            ElseIf Left$(Name, 1) = "'" Then
                sResult = "'" + sResult
            End If

            TempName = sResult

        End If

    End If

    Dim oName As Name
    Set oName = oWB.Names(TempName)              ' Names is a property of Workbook

    sResult = Mid$(oName.RefersTo, 2)            ' remove initial equals-sign (=)
    If Choice = 0 Then
        NameRefersTo = sResult
    Else
        If Left$(sResult, 1) = "'" Then          ' there might be escaped single-quotes ('')

            Dim sSheet As String
            sSheet = Left$(sResult, (InStr(1, Replace(sResult, "''", "xx"), "'!") + 1))
        Else                                     ' no single-quotes; therefore,
            sSheet = Left$(sResult, InStr(1, sResult, "!")) ' no exclamation-point (!) in Worksheet reference
        End If

        If sSheet = vbNullString Then            ' unexpected for a named range
            Dim oWS As Worksheet
            Set oWS = ActiveSheet                ' default Worksheet
        Else
            sSheet = Left$(sSheet, (Len(sSheet) - 1)) ' extract 'xxx[yyy]zzz' from 'xxx[yyy]zzz'!
            sBook = Mid$(sSheet, (InStr(1, sSheet, "[") + 1)) ' extract yyy]zzz' from 'xxx[yyy]zzz'

            If sBook = sSheet Then
                sBook = vbNullString             ' [yyy] was not present
            Else
                sBook = Left$(sBook, InStr(1, sBook, "]")) ' extract yyy] from yyy]zzz'

                If sBook <> vbNullString Then
                    sBook = Left$(sBook, (Len(sBook) - 1)) ' extract yyy from yyy]
                    Set oWB = Workbooks(sBook)   ' referenced Workbook (error if not open)
                    sSheet = Mid$(sSheet, (InStr(1, sSheet, "]") + 1)) ' extract zzz' from xxx[yyy]zzz'
                End If

            End If

            sResult = sSheet
            If Left$(sResult, 1) = "'" Then sResult = Mid$(sResult, 2) ' remove any leading/trailing single-quotes (')
            If Right$(sResult, 1) = "'" Then sResult = Left$(sResult, (Len(sResult) - 1))
            sResult = Replace(sResult, "''", "'") ' replace escaped single-quotes ('')
            Set oWS = oWB.Worksheets(sResult)    ' referenced Worksheet (error if not in oWB)

            If sBook <> vbNullString Then
                sSheet = "[" + sBook + "]" + sSheet ' add any Workbook in brackets
                If Right$(sSheet, 1) = "'" Then sSheet = "'" + sSheet ' fix single-quotes
            End If

        End If

        If Choice < -1 Then
            NameRefersTo = oWS.Name              ' unqualified Worksheet name
        ElseIf Choice < 0 Then
            NameRefersTo = sSheet                ' qualified Worksheet reference (including Workbook name)
        Else
            sResult = Mid$(oName.RefersTo, 2)    ' remove initial equals-sign (=)
            sResult = Replace(sResult, (sSheet + "!#REF!,"), vbNullString) ' remove Sheet!#REF!,
            sResult = Replace(sResult, ("," + sSheet + "!#REF!"), vbNullString) ' remove ,Sheet!#REF!

            If Choice = 1 Then
                sResult = Replace(sResult, (sSheet + "!"), vbNullString) ' remove Worksheet reference
                NameRefersTo = oWS.Range(sResult).Address ' error if invalid Range
            Else
                NameRefersTo = sResult           ' retain Worksheet reference
                sResult = oWS.Range(NameRefersTo).Address ' error if invalid Range
            End If

        End If

    End If

    Exit Function                                ' success (no error)

Return_Error:
    NameRefersTo = CVErr(xlErrValue)             ' #VALUE! (Error 2015)

End Function

Private Function IsProtected( _
        Optional ByVal Choice As Variant = 0, _
        Optional ByVal Target As Range = Nothing _
        ) As Variant
    '
    '   User-Defined Function (UDF) to return the protection status (True or False) of Target's Worksheet or Workbook
    '       see https://excelribbon.tips.net/T009639_Visually_Showing_a_Protection_Status.html
    '   Default Choice is 0 (or "contents")
    '       Choice =  3 or "scenarios"   return True if the Worksheet's scenarios are protected
    '       Choice =  2 or "interface"   return True if the Worksheet's user interface is protected (but not its macros)
    '       Choice =  1 or "shapes"      return True if the Worksheet's shapes are protected
    '       Choice =  0 or "contents"    return True if the Worksheet's contents are protected (this is the default Choice)
    '       Choice = -1 or "sheets"      return True if the order of the Workbook's sheets are protected
    '       Choice = -2 or "windows"     return True if the Workbook's windows are protected
    '   Default Target is the cell referencing this function (error if referenced in a VBA statement);
    '       otherwise, a Worksheet cell's address (like $A$1) or Range such as Range("$A$1")
    '
    '   Feb 2019 by J. Woolley
    '
    Application.Volatile

    Dim TempTarget As Range
    If Target Is Nothing Then Set TempTarget = Application.ThisCell

    Dim TempChoice As Variant
    If Not IsNumeric(Choice) Then TempChoice = LCase$(Choice)

    Select Case TempChoice
    Case 3, "scenarios": IsProtected = TempTarget.Parent.ProtectScenarios
    Case 2, "interface": IsProtected = TempTarget.Parent.ProtectionMode
    Case 1, "shapes": IsProtected = TempTarget.Parent.ProtectDrawingObjects
    Case 0, "contents": IsProtected = TempTarget.Parent.ProtectContents
    Case -1, "sheets": IsProtected = TempTarget.Parent.Parent.ProtectStructure
    Case -2, "windows": IsProtected = TempTarget.Parent.Parent.ProtectWindows
    Case Else: IsProtected = CVErr(xlErrValue)   ' #VALUE! (Error 2015)
    End Select

End Function

Private Function ProgressBar_Text( _
        ByVal sLabel As String, _
        ByVal nCurrent As Long, _
        ByVal nTotal As Long _
        ) As String
    '
    '   Prepare a text string indicating progress in processing sLabel item nCurrent of nTotal items
    '   If sLabel is blank, the bar will be wider without any text
    '
    '   Jul 2018 by J. Woolley
    '   Feb 2019 by J. Woolley
    '

    Dim k As Long
    k = 50

    If sLabel = vbNullString Then k = k * 2

    Dim n As Long
    n = Round(k * (nCurrent - 1) / nTotal, 0)

    ProgressBar_Text = "|" + String(n, "|") + String((k - n), ".") + "|"

    If sLabel <> vbNullString Then ProgressBar_Text = ProgressBar_Text + "  " + sLabel + CStr(nCurrent) + " of " + CStr(nTotal)

End Function

Private Sub UnhideNames()
    '
    '   Macro to unhide all defined names
    '
    '   Feb 2019 by J. Woolley
    '

    Const myName = "UnhideNames"

    Dim Msg As String
    Msg = _
        "Do you want to unhide all defined names (named ranges) in ALL WORKSHEETS of the ACTIVE WORKBOOK " + _
        "or only in the ACTIVE SHEET?" + _
        vbNewLine + vbNewLine + _
        "Click Yes to unhide all names in ALL WORKSHEETS." + _
        vbNewLine + _
        "Click No to unhide all names in the ACTIVE SHEET." + _
        vbNewLine + _
        "Or click Cancel to quit."

    Dim ans As Variant
    ans = MsgBox(Msg, (vbYesNoCancel + vbQuestion), myName)

    If ans = vbYes Then
        Dim bBook As Boolean
        bBook = True

        Dim oThis As Object
        Set oThis = ActiveWorkbook.Names
    ElseIf ans = vbNo Then
        Msg = "Do you want to unhide all names in the ACTIVE SHEET?"
        ans = MsgBox(Msg, (vbYesNo + vbQuestion), myName)
        If ans <> vbYes Then Exit Sub
        bBook = False
        Set oThis = ActiveWorkbook.ActiveSheet.Names
    Else
        Exit Sub
    End If

    Dim gItems As Long
    gItems = oThis.Count

    Dim gHidden As Long
    gHidden = 0

    Dim oName As Name
    For Each oName In oThis
        If Not oName.Visible Then
            gHidden = gHidden + 1
            oName.Visible = True
        End If
    Next oName

    Msg = _
        CStr(gHidden) + " of " + CStr(gItems) + _
        " names in the ACTIVE " + IIf(bBook, "WORKBOOK", "SHEET") + _
        " were hidden. " + _
        "All are visible now in Name Manager (Ctrl+F3) and Go To (F5)."

    MsgBox Msg, vbOKOnly, myName

End Sub

Public Sub NameRefersTo_Register()
    '
    '   Private Sub to register Public Function NameRefersTo
    '   Run this once manually (F5), but not while enabled as Add-In; repeat if there are changes
    '
    '   Feb 2019 by J. Woolley
    '

    Dim sName As String
    sName = "NameRefersTo"

    Dim sDesc As String
    sDesc = "Return the RefersTo property of a defined name (named range) as text."

    Dim nCNbr As Long
    nCNbr = 14                                   ' category name is "User Defined"

    Dim sArgDesc(1 To 2) As String
    sArgDesc(1) = "A defined name (named range) as text"

    sArgDesc(2) = _
                "-2 = unqualified Sheet, -1 = qualified Sheet," + _
                vbNewLine + _
                "0 = RefersTo (default)," + _
                vbNewLine + _
                "1 = no Sheet! or #REF!, 2 = no Sheet!#REF!"

    Dim Msg As String
    Msg = _
        "Do you want to register Function " + _
        sName + "?" + _
        vbNewLine + _
        "(This only needs to be done once, unless there are changes.)"

    Dim ans As Long
    ans = MsgBox(Msg, (vbYesNo + vbQuestion), "Register Function")

    If ans <> vbYes Then Exit Sub

    Application.MacroOptions _
        Macro:="Personal.xlsb!" & sName, _
        Description:=sDesc, _
        Category:=nCNbr, _
        ArgumentDescriptions:=sArgDesc
        
End Sub

Public Sub IsProtected_Register()
    '
    '   Private Sub to register Public Function IsProtected
    '   Run this once manually (F5), but not while enabled as Add-In; repeat if there are changes
    '
    '   Feb 2019 by J. Woolley
    '

    Dim sName As String
    sName = "IsProtected"

    Dim sDesc As String
    sDesc = "Return the protection status (TRUE or FALSE) of Target's Worksheet or Workbook."

    Dim nCNbr As Long
    nCNbr = 14                                   ' category name is "User Defined"

    Dim sArgDesc(1 To 2) As String
    sArgDesc(1) = _
                "Optional number 0 (default), 1, 2, 3, -1, -2 or equivalent text:" + _
                vbNewLine + _
                "Worksheet ""Contents"" (default), ""Shapes"", ""Interface"", ""Scenarios""" + _
                                                                             vbNewLine + _
                                                                             "Workbook ""Sheets"" or ""Windows"""

    sArgDesc(2) = _
                "Optional cell address; for example, $A$1." + _
                vbNewLine + _
                "Default is the cell referencing this function."

    Dim Msg As String
    Msg = _
        "Do you want to register Function " + sName + "?" + _
        vbNewLine + _
        "(This only needs to be done once, unless there are changes.)"

    Dim ans As Long
    ans = MsgBox(Msg, (vbYesNo + vbQuestion), "Register Function")

    If ans <> vbYes Then Exit Sub

    Application.MacroOptions _
        Macro:="Personal.xlsb!" & sName, _
        Description:=sDesc, _
        Category:=nCNbr, _
        ArgumentDescriptions:=sArgDesc

End Sub


