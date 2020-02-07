Attribute VB_Name = "CommonRoutines"
Option Explicit
' Changes
' 9/15/19:
'       Deleted Private Function TestDrive
'       Added TurnOnffAutomaticProcessing
'       Added TurnOnAutomaticProcessing
'       Added TryFindTableInWorksheet
' 9/18/19
'   Changed TryReadTable to TryCopyRangeToArray
'   Added TryCopyTableToArray
' 9/20/19
'   Changed TurnOn and TurnOffAutomaticProcessing to functions
'       so they don't show up in the list of executable routines
'   Reworked TryReadTable to TryCopyRangeToArray
' 9/22/19
'   Added unfreeze to ConvertDataToTable
' 9/27/19
'   Added CheckInRange
' 9/28/19
'   Deleted extraneous code from CheckInRange
' 9/29/19
'   Changed TryCopyRangeToArray VisibleOnly to optional
'   Changed Split calculations in TryCopyRangeToArray
'   Added Application.StatusBar=False to TurnOnAutomaticProcessing to clear the status bar
'   Added TryCopyTableToArrayWithMapping
'   Modified GetASheet to not delete a sheet with a codename other that "Sheet*"
' 9/30/19
'   Updated GetASheet. Wasn't properly handling a new workbook/worksheet.
' 10/1/19
'   Changed TurnOnAutomaticProcessing enableevents to true
' 10/10/19
'   Added CleanTwoDecimalData, CleanToLeftAlignment
' 10/11/19
'   Added ReportError for exceptions
'    Fixed error checking in TryGetFilePath and TryGetFolderPath
' 10/20
'   Added debug information to the RaiseError calls
' 10/25
'   Converted ConvertDataToTable to a function that returns the table

Private Const Module_Name As String = "CommonRoutines."

Public Function CheckStringInRange( _
       ByVal TryString As String, _
       ByVal TryRange As Range _
       ) As Boolean

    Const RoutineName As String = Module_Name & "CheckStringInRange"
    On Error GoTo ErrorHandler

    ' Assume success
    CheckStringInRange = True

    Dim IndexLocation As Long

    IndexLocation = Application.WorksheetFunction.Match(TryString, TryRange, 0)

    If IndexLocation = 0 Then
        ' TryString not found in TryRange
        CheckStringInRange = False
        GoTo Done
    Else
        ' TryString found in TryRange
    End If

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Search String", TryString, _
                "Range", TryRange.Address
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CheckStringInRange

Public Function CheckNameInCollection( _
       ByVal Key As String, _
       ByVal Coll As Object _
       ) As Boolean

    Const RoutineName As String = Module_Name & "CheckNameInCollection"
    On Error GoTo ErrorHandler

    Dim Element As Object

    Dim ErrorNumber As Long
    On Error Resume Next
    Set Element = Coll(Key)
    ErrorNumber = Err.Number
    On Error GoTo 0
    CheckNameInCollection = (ErrorNumber = 0)

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Search String", Key
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' CheckNameInCollection

Public Function TryGetFilePath( _
       ByVal FileType As String, _
       ByVal FileSuffix As String, _
       ByVal FileTitle As String, _
       ByRef FilePath As String _
       ) As Boolean

    ' This routine asks the user for a file
    ' The offered files are limited to the FileSuffix
    ' Returns the file's path
    
    Const RoutineName As String = Module_Name & "TryGetFilePath"
    On Error GoTo ErrorHandler
    
    ' Assume success
    TryGetFilePath = True
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add FileType, FileSuffix
        .Title = FileTitle
        
        Dim ReturnValue As Variant
        ReturnValue = .Show
        
        If ReturnValue <> 0 Then
            If .SelectedItems(1) = vbNullString Then
                TryGetFilePath = False
            Else
                FilePath = .SelectedItems(1)
            End If
        Else
            TryGetFilePath = False
        End If
        
    End With
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "File Type", FileType, _
                "File Suffix", FileSuffix, _
                "File Title", FileTitle, _
                "File Path", FilePath
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryGetFilePath

Public Function TryGetFolderPath( _
       ByVal InitialFolder As String, _
       ByRef FolderPath As String _
       ) As Boolean

    ' This routine asks the user for a folder
    ' Set the initial folder in Fldr
    ' Returns the folder path
    
    Const RoutineName As String = Module_Name & "TryGetFolderPath"
    On Error GoTo ErrorHandler
    
    TryGetFolderPath = True
    
    Dim FSO As FileSystemObject
    Set FSO = New Scripting.FileSystemObject

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = InitialFolder
        .Title = "Select the 3T Folder"
        
        Dim ReturnValue As Variant
        ReturnValue = .Show
        
        If ReturnValue <> 0 Then
            If .SelectedItems.Count <> 1 Then
                TryGetFolderPath = False
            Else
                FolderPath = .SelectedItems(1)
            End If
        Else
            TryGetFolderPath = False
        End If
    End With
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Initial Folder", InitialFolder, _
                "Folder Path", FolderPath
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryGetFolderPath

Public Function TryGetFilesInFolder( _
       ByVal FolderPath As String, _
       ByRef FileList As Variant _
       ) As Boolean

    Const RoutineName As String = Module_Name & "TryGetFilesInFolder"
    On Error GoTo ErrorHandler
    
    TryGetFilesInFolder = True
    
    Dim FSO As FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    Dim Fldr As Folder
    Set Fldr = FSO.GetFolder(FolderPath)
    
    Dim FileObject As Files
    Set FileObject = Fldr.Files

    If FileObject.Count = 0 Then
        TryGetFilesInFolder = False
        GoTo Done
    End If

    Dim Ary As Variant
    ReDim Ary(1 To FileObject.Count)
    
    Dim OneFile As Variant
    Dim I As Long
    
    I = 1
    For Each OneFile In FileObject
        Ary(I) = OneFile.Name
        I = I + 1
    Next

    FileList = Ary
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Folder Path", FolderPath
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryGetFilesInFolder

Public Function BuildFullTracePath( _
       ByVal FileName As String, _
       Optional ByVal FilePath As String = vbNullString _
       ) As String

   ' This routine builds a full pathname from FileName and FilePath
    ' if FilePath = vbNullString, uses the ActiveWorkbook's path
    ' Returns a string with the full path

    Dim pFileName As String
    If Right$(FileName, 4) <> ".txt" Then
        pFileName = FileName & ".txt"
    Else
        pFileName = FileName
    End If

    Dim pFilePath As String
    If FilePath = vbNullString Then
        pFilePath = ThisWorkbook.Path
    Else
        pFilePath = FilePath
    End If

    BuildFullTracePath = pFilePath & Application.PathSeparator & pFileName

End Function ' BuildFullTracePath

Public Function DesktopFolder() As String

    ' This routine returns the full pathname to the Windows desktop folder

    Const RoutineName As String = Module_Name & "DesktopFolder"
    On Error GoTo ErrorHandler

    Dim objSFolders As Object
    Set objSFolders = CreateObject("WScript.Shell").specialfolders
    DesktopFolder = objSFolders("desktop")

Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' DesktopFolder

Public Function ConvertColumnLetterToNumber(ByVal ColumnLetter As String) As Long
    'PURPOSE: Convert a given letter into it's corresponding Numeric Reference
    'SOURCE: https://www.thespreadsheetguru.com/the-code-vault/vba-code-to-convert-column-number-to-letter-or-letter-to-number

    ConvertColumnLetterToNumber = Range(ColumnLetter & 1).Column

End Function ' ConvertColumnLetterToNumber

Public Function ConvertColumnNumberToLetter(ByVal ColumnNumber As Long) As String
    'PURPOSE: Convert a given number into it's corresponding Letter Reference
    'SOURCE: https://www.thespreadsheetguru.com/the-code-vault/vba-code-to-convert-column-number-to-letter-or-letter-to-number

    ConvertColumnNumberToLetter = Split(Cells(1, ColumnNumber).Address, "$")(1)

End Function ' ConvertColumnNumberToLetter

Public Sub ClearTable(ByVal LstObj As ListObject)
    LstObj.Parent.Activate
    On Error Resume Next
    LstObj.Parent.ShowAllData
    On Error GoTo 0

    If LstObj.ListRows.Count > 1 Then
        LstObj.DataBodyRange.Delete
    ElseIf LstObj.ListRows.Count > 0 Then
        LstObj.DataBodyRange.Clear
    End If
    
    LstObj.Parent.Cells.ClearFormats

End Sub ' ClearTable

Public Function FindLastRow(ByVal ColLetter As String, ByVal RowNumber As Long, _
                            ByVal Sheet As Worksheet) As Long
    Dim RegionRow As Long: RegionRow = Sheet.Range(ColLetter & RowNumber).CurrentRegion.Rows.Count
    Dim ColumnRow As Long: ColumnRow = Sheet.Range(ColLetter & Sheet.Rows.Count).End(xlUp).Row
    Dim ColumnNumber As Long: ColumnNumber = Sheet.Range(ColLetter & 1).Column
    Dim I As Long
    Dim CurrentCell As Range

    If RegionRow = ColumnRow Then
        FindLastRow = ColumnRow
    Else
        For I = Application.WorksheetFunction.Max(ColumnRow, RegionRow) To Application.WorksheetFunction.Min(ColumnRow, RegionRow) Step -1
            Set CurrentCell = Sheet.Cells(I, ColumnNumber)
            If Not IsEmpty(CurrentCell) Then
                FindLastRow = I
                Exit For
            End If
        Next I
    End If
End Function ' FindLastRow

Public Function FindLastColumn(ByVal RowNumber As Long, _
                               ByVal Sheet As Worksheet) As Long

    FindLastColumn = Sheet.Cells(RowNumber, Sheet.Columns.Count).End(xlToLeft).Column
End Function ' FindLastColumn

Public Function ConvertDataToTable( _
       ByVal Wksht As Worksheet, _
       ByVal TableName As String _
       ) As ListObject

    ' This routine converts the data on a worksheet to a table
    ' It assumes the data starts in $A$1
    
    Const RoutineName As String = Module_Name & "ConvertDataToTable"
    On Error GoTo ErrorHandler
    
    Wksht.Activate
    ActiveWindow.FreezePanes = False
    
    Wksht.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
    
    Dim LastRow As Long
    LastRow = FindLastRow("A", 1, Wksht)
    
    Dim LastColumnLetter As String
    LastColumnLetter = ConvertColumnNumberToLetter(FindLastColumn(1, Wksht))
    
    Wksht.ListObjects.Add(xlSrcRange, Range("$A$1:$" & LastColumnLetter & "$" & LastRow), , xlYes).Name = TableName
    
    Set ConvertDataToTable = Wksht.ListObjects(TableName)
    
    Wksht.Cells.EntireColumn.AutoFit
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Worksheet", Wksht.Name, _
                "Table Name", TableName
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' ConvertDataToTable

Public Function GetASheet( _
       ByVal Wkbk As Workbook, _
       ByVal SheetName As String _
       ) As Worksheet

    ' This routine
    '   Checks to see if Wksht exists in Wkbk
    '   Creates a sheet named SheetName in Wkbk if Wksht does not exist
    '   The created sheet is created as the last sheet in the wkbk
    ' Returns an existing or new worksheet
    
    Const RoutineName As String = Module_Name & "GetASheet"
    On Error GoTo ErrorHandler
    
    ' Determine if sheet already exists
    Dim ThisSheet As Worksheet
    Dim ErrorNumber As Long
    On Error Resume Next
    Set ThisSheet = Wkbk.Sheets(SheetName)
    ErrorNumber = Err.Number
    On Error GoTo ErrorHandler
            
    If ErrorNumber <> 0 Then
        ' The named sheet does not exist in wkbk
        Set ThisSheet = Wkbk.Sheets.Add(after:=Wkbk.Worksheets(Wkbk.Worksheets.Count))
        ThisSheet.Name = SheetName
    Else
        ' The named sheet does exist in wkbk
        Dim Previous As Boolean
        
        ' Do not delete a sheet that already has a changed codename
        If Left$(Wkbk.Worksheets(SheetName).CodeName, 5) = "Sheet" Then
            Previous = Application.DisplayAlerts
            Application.DisplayAlerts = False
            Wkbk.Worksheets(SheetName).Delete
            Application.DisplayAlerts = Previous
            
            Set ThisSheet = Wkbk.Sheets.Add(after:=Wkbk.Worksheets(Wkbk.Worksheets.Count))
            ThisSheet.Name = SheetName
        Else
            Set ThisSheet = Wkbk.Worksheets(SheetName)
            ThisSheet.Cells.ClearContents
        End If
    End If
    
    Set GetASheet = ThisSheet
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Workbook", Wkbk.Name, _
                "Worksheet", SheetName
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' GetASheet

Public Function TryCopyRangeToArray( _
       ByVal Rng As Range, _
       ByRef Result As Variant, _
       Optional ByVal VisibleOnly As Boolean = False _
       ) As Boolean

    ' This routine converts the data in Rng to a 2D array
    ' If VisibleOnly = True, then the hidden rows of Rng are skipped
    ' Assumption: the first row of the range contains column headers and the data is in the successive rows
    ' Returns the array from the rng
    ' Returns True if successful
    
    Const RoutineName As String = Module_Name & "TryCopyRangeToArray"
    On Error GoTo ErrorHandler
    
    TryCopyRangeToArray = True
    
    Dim SplitArray As Variant
    SplitArray = Split(Rng.Address, "$")
    
    Dim StartRow As Long
    StartRow = Split(SplitArray(2), ":")(0)
    
    Dim StopRow As Long
    StopRow = SplitArray(4)
    
    Dim StartColumn As Long
    StartColumn = ConvertColumnLetterToNumber(SplitArray(1))
    
    Dim StopColumn As Long
    StopColumn = ConvertColumnLetterToNumber(SplitArray(3))
    
    Dim pAry() As Variant
    ReDim pAry(StopRow - StartRow, StopColumn - StartColumn)
    
    Dim TableArray As Variant
    TableArray = Rng
    
    If VisibleOnly Then
        ' Exclude the hidden rows
        Dim NumberOfRows As Long
        NumberOfRows = 0
        
        Dim IndividualRange As Range
        For Each IndividualRange In Rng.Areas
            SplitArray = Split(IndividualRange, "$")

            StartRow = SplitArray(2)

            StopRow = SplitArray(4)

            Dim I As Long
            For I = StartRow To StopRow
                Dim J As Long
                For J = StartColumn To StopColumn
                    pAry(NumberOfRows, J - 1) = TableArray(I, J)
                Next J

                NumberOfRows = NumberOfRows + 1
            Next I
        Next IndividualRange
        
        Dim tAry() As Variant
        ReDim tAry(NumberOfRows - 1, StopColumn - StartColumn)
        
        For I = 0 To NumberOfRows - 1
            For J = 0 To StopColumn - StartColumn
                tAry(I, J) = pAry(I, J)
            Next J
        Next I
        Result = tAry
    Else
        ' Including any hidden rows
        Result = Rng
    End If
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Search Range", Rng.Address, _
                "Visible Only?", VisibleOnly
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyRangeToArray

Public Function TryCopyTableToArray( _
       ByVal Tbl As ListObject, _
       ByVal VisibleOnly As Boolean, _
       ByRef Result As Variant _
       ) As Boolean

    ' This routine copies a table to an array
    
    Const RoutineName As String = Module_Name & "TryCopyTableToArray"
    On Error GoTo ErrorHandler
    
    If TryCopyRangeToArray(Tbl.Range, Result, VisibleOnly) Then
        TryCopyTableToArray = True
    Else
        TryCopyTableToArray = False
        ReportError "Error copying table to array", _
                    "Routine", RoutineName
        GoTo Done
    End If

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Table", Tbl.Name, _
                "Visible Only?", VisibleOnly
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyTableToArray

Public Function TryFindCellInSheet( _
       ByVal Target As String, _
       ByVal Wksht As Worksheet, _
       ByRef Location As String _
       ) As Boolean

    ' This routine searches for Target in WkSht
    ' Returns the cell address (e.g., $A$1) and True if it finds it
    ' Returns False if Target is not found in the first 1000 rows and 1000 columns of WkSht
    
    Const RoutineName As String = Module_Name & "TryFindCellInSheet"
    On Error GoTo ErrorHandler
    
    TryFindCellInSheet = True
    
    Const Limit As Long = 1000
    
    Dim Ary As Variant
    Ary = Wksht.Range("$A$1:" & ConvertColumnNumberToLetter(Limit) & Limit)
    
    Dim I As Long
    Dim J As Long
    For I = 1 To Limit
        For J = 1 To Limit
            If Ary(I, J) = Target Then
                Location = "$" & ConvertColumnNumberToLetter(J) & "$" & I
                GoTo Done
            End If
        Next J
    Next I
    
    TryFindCellInSheet = False
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Search String", Target, _
                "Worksheet", Wksht.Name, _
                "Location", Location
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryFindCellInSheet

Public Function TurnOffAutomaticProcessing()

    ' This routine turns off all the automatic processing that slows things down
    
    Const RoutineName As String = Module_Name & "SubTemplate"
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' SubTemplate

Public Function TurnOnAutomaticProcessing()

    ' This routine turns on all the automatic processing that slows things down
    ' Reverses the things that were turned off in TurnOffAutomaticProcessing
    
    Const RoutineName As String = Module_Name & "TurnOnAutomaticProcessing"
    On Error GoTo ErrorHandler
    
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TurnOnAutomaticProcessing

Public Function TryFindTableInWorksheet( _
       ByVal Wksht As Worksheet, _
       ByVal TableName As String, _
       ByRef Tbl As ListObject _
       ) As Boolean

    ' This routine checks to see if TableName appears in Wksht
    ' Returns true if it does
    ' Returns the table if it does
    
    Const RoutineName As String = Module_Name & "TryFindTableInWorksheet"
    On Error GoTo ErrorHandler
    
    TryFindTableInWorksheet = True
    
    Dim ErrorNumber As Long
    On Error Resume Next
    Set Tbl = Wksht.ListObjects(TableName)
    ErrorNumber = Err.Number
    On Error GoTo ErrorHandler
            
    If ErrorNumber <> 0 Then
        ReportError "Table not found in worksheet", _
                    "Routine", RoutineName, _
                    "Worksheet", Wksht.Name, _
                    "Table", TableName
        TryFindTableInWorksheet = False
        GoTo Done
    Else
        ' Success; return the table
    End If

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Worksheet", Wksht.Name, _
                "Table", TableName
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryFindTableInWorksheet

Public Function CheckInRange( _
       ByVal Range1 As Range, _
       ByVal Range2 As Range _
       ) As Boolean
     
    ' returns True if Range1 is within Range2
    
    CheckInRange = Not (Application.Intersect(Range1, Range2) Is Nothing)
    
Done:
    Exit Function
    
ErrorHandler:
    ' Application.Intersect raises an error if the ranges are on different worksheets
    CheckInRange = False
    On Error GoTo 0
End Function ' CheckInRange

Public Function TryCopyTableToArrayWithMapping( _
       ByVal OldTbl As ListObject, _
       ByVal Mapping As Variant, _
       ByRef NewAry As Variant _
       ) As Boolean
        
    ' This routine copies all the data from one table and rearranges it based on the Mapping array
    ' Mapping is an array of arrays that starts at 0 in all dimensions

    Const RoutineName As String = Module_Name & "TryCopyTableToArrayWithMapping"
    On Error GoTo ErrorHandler
    
    TryCopyTableToArrayWithMapping = True
    
    Dim Ary As Variant
    Ary = OldTbl.Parent.Range(OldTbl.Range.Address)
    
    ReDim NewAry(1 To UBound(Ary, 1) - 1, 1 To UBound(Mapping, 1) + 1)
    
    Dim I As Long
    Dim J As Long
    Dim OldCol As Long
    
    ' Loop through the rows of the mapping array
    For I = 0 To UBound(Mapping, 1)
        ' Loop across the columns of each row of the maping array
        For J = 1 To UBound(Ary, 2)
            If Ary(1, J) = Mapping(I)(0) Then
                OldCol = J
                Exit For
            End If
        Next J
        
        Dim NewCol As Long
        For J = 0 To UBound(Mapping, 1)
            If Ary(1, OldCol) = Mapping(J)(0) Then
                NewCol = J + 1
                Exit For
            End If
        Next J
        
        For J = 1 To UBound(Ary, 1) - 1
            NewAry(J, NewCol) = Ary(J + 1, OldCol)
        Next J
    Next I
    
Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Old Table", OldTbl.Name
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyTableToArrayWithMapping

Public Sub CleanTwoDecimalData( _
       ByVal Tbl As ListObject, _
       ByVal ColumnNumber As Long)

    ' This routine cleans up the formatting for decimals
    
    Const RoutineName As String = Module_Name & "CleanTwoDecimalData"
    On Error GoTo ErrorHandler
    
    Dim Rng As Range
    Set Rng = Tbl.ListColumns(ColumnNumber).Range
    Rng.Style = "Comma"
    
Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Table", Tbl.Name
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' CleanTwoDecimalData

Public Sub CleanToLeftAlignment( _
       ByVal Tbl As ListObject, _
       ByVal ColumnNumber As Long)

    ' This routine sets the column to left alignment

    Const RoutineName As String = Module_Name & "CleanToLeftAlignment"
    On Error GoTo ErrorHandler

    Dim Rng As Range
    Set Rng = Tbl.ListColumns(ColumnNumber).Range
    Rng.HorizontalAlignment = xlLeft

Done:
    Exit Sub
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description, _
                "Table", Tbl.Name
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub ' CleanToLeftAlignment




