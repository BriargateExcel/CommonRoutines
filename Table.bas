Attribute VB_Name = "Table"
Option Explicit

Private Const Module_Name As String = "Table."

Public Function TryCopyDictionaryToTable( _
    ByVal TableType As iTable, _
    Optional ByVal Dict As Dictionary = Nothing, _
    Optional ByVal Tbl As ListObject = Nothing, _
    Optional ByVal Rng As Range = Nothing, _
    Optional ByVal TableName As String = vbNullString _
    ) As Boolean

    ' This routine copies a dictionary to a table
    ' If Dict is nothing then use pLocalDict dictionary
    ' If Tbl is nothing then build a table using Rng and TableName
    ' if Tbl and Rng are both Nothing then use the main table

    Const RoutineName As String = Module_Name & "CopyDictionaryToTable"
    On Error GoTo ErrorHandler
    
    TryCopyDictionaryToTable = True

    Dim ThisDict As Dictionary
    If Dict Is Nothing Then
        If Not TableType.Initialized Then TableType.Initialize
        Set ThisDict = TableType.LocalDictionary
    Else
        If Dict.Count = 0 Then
            TryCopyDictionaryToTable = False
            GoTo Done
        End If
        Set ThisDict = Dict
    End If

    Dim ThisTbl As ListObject
    If Tbl Is Nothing Then
        If Rng Is Nothing Then
            Set ThisTbl = TableType.LocalTable
        Else
            If TableName = vbNullString Then
                ReportError "Need to provide a table name", "Routine", RoutineName
                TryCopyDictionaryToTable = False
                GoTo Done
            Else
                Set ThisTbl = Rng.Parent.ListObjects.Add(xlSrcRange, _
                    Range(Cells(1, 1), Cells(2, TableType.HeaderWidth)), , xlYes)
                ThisTbl.Name = TableName
            End If
        End If
    Else
        Set ThisTbl = TableType.LocalTable
    End If
    
    Dim AddressPieces As Variant
    AddressPieces = Split(ThisTbl.HeaderRowRange.Address, ":")
    
    Dim ThisRng As Range
    Set ThisRng = ThisTbl.Parent.Range(AddressPieces(0))

    ThisRng.Resize(1, TableType.HeaderWidth) = TableType.Headers
    
    ClearTable ThisTbl

    Dim Ary As Variant
    ReDim Ary(1 To ThisDict.Count, 1 To TableType.HeaderWidth)

    If TableType.TryCopyDictionaryToArray(ThisDict, Ary) Then
        ' Success; do nothing
    Else
        ReportError "Error copying dictionary to array", "Routine", RoutineName
        TryCopyDictionaryToTable = False
        GoTo Done
    End If
    
    ' Format the worksheet
    TableType.FormatArrayAndWorksheet Ary, ThisTbl
    
    ' move to DatabodyRange
    Set ThisRng = ThisRng.Offset(1, 0)
    ThisRng.Resize(UBound(Ary, 1), TableType.HeaderWidth) = Ary
    ThisRng.Resize(UBound(Ary, 1), TableType.HeaderWidth) = Ary ' Seems to be needed to get the column formatting right

    ThisRng.Parent.Cells.EntireColumn.AutoFit

    ThisRng.Parent.Activate
    ActiveWindow.FreezePanes = False

    ThisRng.Parent.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyDictionaryToTable

Public Function TryCopyTableToDictionary( _
    ByVal TableType As iTable, _
    ByVal Tbl As ListObject, _
    Optional ByRef Dict As Dictionary _
    ) As Boolean

    ' Copies a table to a dictionary

    Const RoutineName As String = Module_Name & "TryCopyTableToDictionary"
    On Error GoTo ErrorHandler

    TryCopyTableToDictionary = True

    Dim Ary As Variant
    On Error Resume Next
    Ary = Tbl.DataBodyRange
    If Err.Number <> 0 Then
        ReportError "The " & TableType.LocalName & " table is empty", "Routine", RoutineName
        TryCopyTableToDictionary = False
        GoTo Done
    End If
    Err.Clear
    On Error GoTo ErrorHandler

    Dim ThisDict As Dictionary
    If Dict Is Nothing Then
        Set ThisDict = New Dictionary
    Else
        Set ThisDict = TableType.LocalDictionary
    End If

    If TableType.TryCopyArrayToDictionary(Ary, ThisDict) Then
        ' Success; do nothing
    Else
        ReportError "Error loading dictionary", "Routine", RoutineName
    End If

    Set Dict = ThisDict

Done:
    Exit Function
ErrorHandler:
    ReportError "Exception raised", _
                "Routine", RoutineName, _
                "Error Number", Err.Number, _
                "Error Description", Err.Description
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function ' TryCopyTableToDictionary


