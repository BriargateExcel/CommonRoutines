# Common Routines

Stuff that's useful across many projects

## CommonRoutines Standard Module

```
Public Function CheckStringInRange( _
   ByVal TryString As String, _
   ByVal TryRange As Range _
   ) As Boolean
```

```
Public Function CheckNameInCollection( _
       ByVal Key As String, _
       ByVal Coll As Object _
       ) As Boolean
```

```
Public Function TryGetFilePath( _
       ByVal FileType As String, _
       ByVal FileSuffix As String, _
       ByVal FileTitle As String, _
       ByRef FilePath As String _
       ) As Boolean
```

```
Public Function TryGetFolderPath( _
    ByVal InitialFolder As String, _
    ByRef FolderPath As String _
    ) As Boolean
```

```
Public Function TryGetFilesInFolder( _
       ByVal FolderPath As String, _
       ByRef FileList As Variant _
       ) As Boolean
```

```
Public Function BuildFullTracePath( _
       ByVal Filename As String, _
       Optional ByVal FilePath As String = vbNullString _
       ) As String
```

```
Public Function DesktopFolder() As String
```

```
Public Function ConvertColumnLetterToNumber(ByVal ColumnLetter As String) As Long
```

```
Public Function ConvertColumnNumberToLetter(ByVal ColumnNumber As Long) As String
```

```
Public Sub ClearTable(ByVal LstObj As ListObject)
```

```
Public Function FindLastRow(ByVal ColLetter As String, ByVal RowNumber As Long, _
                            ByVal Sheet As Worksheet) As Long
```

```
Public Function FindLastColumn(ByVal RowNumber As Long, _
                               ByVal Sheet As Worksheet) As Long
```

```
Public Sub ConvertDataToTable( _
       ByVal Wksht As Worksheet, _
       ByVal TableName As String)
```

```
Public Function GetASheet( _
        ByVal Wkbk As Workbook, _
        ByVal SheetName As String _
        ) As Worksheet
```

```
Public Function TryReadTable( _
       ByVal Tbl As ListObject, _
       ByVal VisibleOnly As Boolean, _
       ByRef Result As Variant _
       ) As Boolean
```






## CSVFormatClass Class Module

```
Public Sub SetOneFormat( _
       ByVal ExpectedHeader As String, _
       ByVal ColumnTitle As String, _
       ByVal ColumnOrder As Long, _
       ByVal DataFormat As String, _
       ByVal FieldOrder As Long)
```

```
Public Property Let HeaderFound(ByVal FoundValue As Boolean): pHeaderFound = FoundValue: End Property

Public Property Get HeaderFound() As Boolean: HeaderFound = pHeaderFound: End Property

Public Property Get DataFormat() As String: DataFormat = pDataFormat: End Property

Public Property Get ColumnOrder() As Long: ColumnOrder = pColumnOrder: End Property

Public Property Get Header() As String: Header = pExpectedHeader: End Property

Public Property Get ColumnTitle() As String: ColumnTitle = pColumnTitle: End Property
```

## CSVHandler Class Module

```
Public Function TryInputCSVFileToTable( _
       ByVal FormatSheet As Worksheet, _
       ByVal FileSelectionBoxTitle As String, _
       ByVal TableName As String, _
       ByVal SheetName As String, _
       ByVal StopIfErrorFound As Boolean, _
       ByVal CSVFilePath As String, _
       ByVal CopyCSVFileToTable As Boolean, _
       ByVal NewWorksheet As Boolean, _
       ByVal UpperLeftCorner As String, _
       ByRef Result As Variant _
       ) As Boolean
```

## ErrorFileClass Class Module

```
Public Sub WriteErrorLine(ByVal ErrorMessage As String)
```

```
Public Sub WriteBlankErrorLines(Optional ByVal NumLines As Long = 1)
```

## ErrorRoutines Standard Module

```
Public Sub RaiseError( _
       ByVal ErrorNo As Long, _
       ByVal Src As String, _
       ByVal Proc As String, _
       ByVal Desc As String, _
       ParamArray Args() As Variant)
```

```
Public Sub DisplayError(ByVal ProcName As String)
```

```
Public Sub ReportError( _
       ByVal ErrMsg As String, _
       ParamArray Args() As Variant)
```

```
Public Function CloseErrorFile() As Boolean
```

