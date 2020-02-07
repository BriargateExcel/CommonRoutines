Attribute VB_Name = "PerformanceRoutines"
Option Explicit
Option Private Module
' Timer comes from https://www.jkp-ads.com/Articles/performanceclass.asp

' Requires PerformanceClass

' How to use:
'Sub MainProgram()
'    Dim cPerf As PerformanceClass
'    ResetPerformance
'    If gbDebug(RoutineName) Then
'        Set cPerf = New PerformanceClass
'        cPerf.SetRoutine RoutineName
'    End If
'    Application.OnTime Now, "ReportPerformance"
'End Sub
'
'Sub Subroutine()
'    Dim cPerf As PerformanceClass
'    If gbDebug(RoutineName) Then
'        Set cPerf = New PerformanceClass
'        cPerf.SetRoutine RoutineName
'    End If
'End Sub

Private Const Module_Name As String = "PerformanceRoutines."

'For PerformanceClass:
Private pPerformanceIndex As Long                ' Index into the PerformanceResults array
Private pPerformanceResults() As Variant         ' The timing data
Private pCallDepth As Long                       ' Measures the depth of the call stack

Private Const PerformanceSheetName As String = "_Performance_"

' Timer comes from https://www.jkp-ads.com/Articles/performanceclass.asp

#If VBA7 Then
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

Public Sub SetPerformanceResults( _
       ByVal RowNum As Long, _
       ByVal ColNum As Long, _
       ByVal NewValue As Variant)

    ' This routine sets a value in Perfromance Results
    
    Const RoutineName As String = Module_Name & "SetPerformanceResults"
    On Error GoTo ErrorHandler
    
    pPerformanceResults(RowNum, ColNum) = NewValue
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' SetPerformanceResults

Public Sub ReDimPerformanceResults( _
       ByVal LowerBound As Long, _
       ByVal UpperBound As Long)

    ' This routine redims PerformanceResults preserving the contents
    
    Const RoutineName As String = Module_Name & "ReDimPerformanceResults"
    On Error GoTo ErrorHandler
    
    ReDim pPerformanceResults(1 To LowerBound, 1 To UpperBound)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ReDimPerformanceResults

Public Sub ReDimPreservePerformanceResults( _
       ByVal LowerBound As Long, _
       ByVal UpperBound As Long)

    ' This routine redims PerformanceResults preserving the contents
    
    Const RoutineName As String = Module_Name & "ReDimPreservePerformanceResults"
    On Error GoTo ErrorHandler
    
    ReDim Preserve pPerformanceResults(1 To LowerBound, 1 To UpperBound)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ReDimPreservePerformanceResults

Public Function PerformanceResultsBounded() As Boolean

    ' This routine returns true if PerformanceResults is bounded
    
    Const RoutineName As String = Module_Name & "PerformanceResultsBounded"
    On Error GoTo ErrorHandler
    
    PerformanceResultsBounded = IsBounded(pPerformanceResults)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' PerformanceResultsBounded

Public Sub DecrementCallDepth()

    ' This routine decrements CallDepth
    
    Const RoutineName As String = Module_Name & "DecrementCallDepth"
    On Error GoTo ErrorHandler
    
    pCallDepth = pCallDepth - 1
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' DecrementCallDepth

Public Sub IncrementCallDepth()

    ' This routine increments CallDepth
    
    Const RoutineName As String = Module_Name & "IncrementCallDepth"
    On Error GoTo ErrorHandler
    
    pCallDepth = pCallDepth + 1
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' IncrementCallDepth

Public Function CallDepth() As Long

    ' This routine returns the value of the PerformanceIndex
    
    Const RoutineName As String = Module_Name & "CallDepth"
    On Error GoTo ErrorHandler
    
    CallDepth = pCallDepth
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' CallDepth

Public Function PerformanceIndex() As Long

    ' This routine returns the value of the PerformanceIndex
    
    Const RoutineName As String = Module_Name & "PerformanceIndex"
    On Error GoTo ErrorHandler
    
    PerformanceIndex = pPerformanceIndex
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' PerformanceIndex

Public Sub IncrementPerformanceIndex()

    ' This routine increments the PerformanceIndex
    
    Const RoutineName As String = Module_Name & "IncrementPerformanceIndex"
    On Error GoTo ErrorHandler
    
    pPerformanceIndex = pPerformanceIndex + 1
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' IncrementPerformanceIndex

Public Sub ResetPerformanceIndex()

    ' This routine set the PerformanceIndex back to 0
    
    Const RoutineName As String = Module_Name & "ResetPerformanceIndex"
    On Error GoTo ErrorHandler
    
    pPerformanceIndex = 0
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ResetPerformanceIndex

Public Function gbDebug( _
       Optional ByVal CallingRoutine As String = vbNullString _
       ) As Boolean

    ' This routine determines what debug actions to take depending on which routine called it
    ' Returns True if we're measuring performance for the calling routine
    
    Const RoutineName As String = Module_Name & "gbDebug"
    On Error GoTo ErrorHandler
    
    gbDebug = False
    Exit Function
        
    Select Case CallingRoutine
    Case "ExcelRainManProject.IDFound"
        gbDebug = False
    Case "TextFileClass.WriteBlankLinesToFile"
        gbDebug = False
    Case "TextFileClass.WriteLineToFile"
        gbDebug = False
    Case "TextFileClass.BuildFullPath"
        gbDebug = False
    Case "TextFileClass.FileExists"
        gbDebug = False
    Case "TextFileClass.OpenFile"
        gbDebug = False
    Case "TextFileClass.WriteBlankLinesToFile"
        gbDebug = False
    Case "TextFileClass.WriteLineToFile"
        gbDebug = False
    Case "TextFileClass.FileExists"
        gbDebug = False
    Case "TextFileClass.OpenFile"
        gbDebug = False
    Case "TextFileClass.WriteBlankLinesToFile"
        gbDebug = False
    Case "TextFileClass.WriteLineToFile"
        gbDebug = False
    Case "ErrorLogClass.ReportError"
        gbDebug = False
    Case "ErrorLogClass.Class_Terminate"
        gbDebug = False
    Case "ErrorLogClass.Class_Terminate"
        gbDebug = False
    Case "ErrorLogClass.CloseErrorLog"
        gbDebug = False
    Case "UtilityRoutines.Initialize"
        gbDebug = False
    Case "TableArrayClass.ReplaceTable"
        gbDebug = False
    Case "TableArrayClass.InitializeTableArray"
        gbDebug = False
    Case "TableArrayClass.SetpBodyToDict"
        gbDebug = False
    Case "TableArrayClass.GetBody"
        gbDebug = False
    Case "TableArrayClass.AddARow"
        gbDebug = False
    Case "TableArrayClass.Found"
        gbDebug = False
    Case "TableArrayClass.GetRowCount"
        gbDebug = False
    Case "TableArrayClass.ColumnNumber"
        gbDebug = False
    Case "UtilityRoutines.WrapUp"
        gbDebug = False
    Case Else
        gbDebug = True
    End Select
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' gbDebug

Public Sub ResetPerformance()
    ResetPerformanceIndex
    ReDim pPerformanceResults(1 To 3, 1 To 1)
End Sub

Public Sub ReportPerformance()
    ' This routine reports the performance of the timed routines
    
    Const RoutineName As String = Module_Name & "ReportPerformance"
    On Error GoTo ErrorHandler
    
    Dim vNewPerf() As Variant
    Dim lRow As Long
    Dim lCol As Long
    ReDim vNewPerf(1 To UBound(pPerformanceResults, 2) + 1, 1 To 3)
    vNewPerf(1, 1) = "Routine"
    vNewPerf(1, 2) = "Started at"
    vNewPerf(1, 3) = "Time taken"
    
    For lRow = 1 To UBound(pPerformanceResults, 2)
        For lCol = 1 To 3
            vNewPerf(lRow + 1, lCol) = pPerformanceResults(lCol, lRow)
        Next
    Next
    
    Dim TestSheet As String
    Dim ErrorNumber As Long
    On Error Resume Next
    TestSheet = ThisWorkbook.Worksheets(PerformanceSheetName).Name
    ErrorNumber = Err.Number
    On Error GoTo ErrorHandler
            
    Dim PerfSheet As Worksheet
    Select Case ErrorNumber
    Case 0
        ' Sheet exists; delete and recreate it
        Set PerfSheet = ThisWorkbook.Worksheets(PerformanceSheetName)
        Application.DisplayAlerts = False
        PerfSheet.Delete
        Application.DisplayAlerts = True
        Set PerfSheet = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets("Items"))
        PerfSheet.Name = PerformanceSheetName
    Case Else
        ' Create the sheet
        Set PerfSheet = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets("Items"))
        PerfSheet.Name = PerformanceSheetName
    End Select
    
    With PerfSheet
        .Range("A1").Resize(UBound(vNewPerf, 1), 3).Value = vNewPerf
        .UsedRange.EntireColumn.AutoFit
    End With
    
    PerfSheet.Range("A1").Select
    Application.CutCopyMode = False
    PerfSheet.ListObjects.Add(xlSrcRange, PerfSheet.UsedRange, , xlYes).Name = "PerformanceTable"
        
    AddPivot PerfSheet
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ReportPerformance

Private Sub AddPivot(ByVal PerfSheet As Worksheet)
    ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=PerfSheet.UsedRange.Address(external:=True), _
        Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:=PerfSheet.Range("$H$3"), _
        TableName:="PerfReport", _
        DefaultVersion:= _
        xlPivotTableVersion14
    
    PerfSheet.Range("$E$1").Select
    
    With PerfSheet.PivotTables("PerfReport")
        With .PivotFields("Routine")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        .AddDataField ActiveSheet.PivotTables("PerfReport").PivotFields("Time taken"), "Average Time Per Call", xlAverage
        .PivotFields("Routine").AutoSort xlDescending, "Average Time Per Call"
        .PivotFields("Average Time Per Call").NumberFormat = "0.000 000"
        
        .AddDataField .PivotFields("Time taken"), "Times Called", xlCount
        .PivotFields("Times Called").NumberFormat = "#,##0"
        
        .AddDataField ActiveSheet.PivotTables("PerfReport").PivotFields("Time taken"), "Total Time", xlSum
        .PivotFields("Total Time").NumberFormat = "0.000 000"
        
        .RowAxisLayout xlTabularRow
        .ColumnGrand = False
        .RowGrand = False
    
    End With

End Sub

Public Function dMicroTimer() As Double
    '-------------------------------------------------------------------------
    ' Procedure : dMicroTimer
    ' Author    : Charles Williams www.decisionmodels.com
    ' Created   : 15-June 2007
    ' Purpose   : High resolution timer
    '             Used for speed optimisation
    '-------------------------------------------------------------------------

    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    dMicroTimer = 0
    If cyFrequency = 0 Then getFrequency cyFrequency
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    If cyFrequency Then dMicroTimer = cyTicks2 / cyFrequency
End Function

Private Function IsBounded(ByVal vArray As Variant) As Boolean
    Dim lTest As Long
    On Error Resume Next
    lTest = UBound(vArray)
    On Error GoTo 0
    IsBounded = (Err.Number = 0)
End Function


