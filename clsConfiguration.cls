VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfiguration"
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

Const Module_Name As String = "clsConfiguration."

' Version 1.0
' Combined the VBAProject and VBAProjectName routines
' Deleted VBAProjectName
' Version 1.0.2
' Added data structure description of pConfig
' Version 1.0.4
' Refactored GetReferences out of MakeConfigFile

Private Const ForReading As Long = 1

Private Const STR_CONFIGKEY_PROJECTNAME             As String = "VBAProject Name"
Private Const STR_CONFIGKEY_BASEPATH                As String = "Base Path"
Private Const STR_CONFIGKEY_MODULEPATHS             As String = "Module Paths"
Private Const STR_CONFIGKEY_REFERENCES              As String = "References"
Private Const STR_CONFIGKEY_REFERENCE_NAME          As String = "Name"
Private Const STR_CONFIGKEY_REFERENCE_DESCRIPTION   As String = "Description"
Private Const STR_CONFIGKEY_REFERENCE_GUID          As String = "GUID"
Private Const STR_CONFIGKEY_REFERENCE_MAJOR         As String = "Major"
Private Const STR_CONFIGKEY_REFERENCE_MINOR         As String = "Minor"
Private Const STR_CONFIGKEY_REFERENCE_PATH          As String = "Path"

Private pProject    As VBProject
Private pConfigFile As String
Private pConfig     As Dictionary
' The pConfig data structure:
' pConfig is a dictionary with 4 elements
' 1. STR_CONFIGKEY_PROJECTNAME is a pConfig dictionary element with a string item, the VBA Project Name
' 2. STR_CONFIGKEY_BASEPATH is a pConfig dictionary element with a string item, base path
' 3. STR_CONFIGKEY_MODULEPATHS is a pConfig dictionary element with a dictionary as its item
'   Each element in the module paths dictionary is a module name with a string item which is the file name for the module
' 4. STR_CONFIGKEY_REFERENCES is pConfig dictionary element with a collection as its item
'   Each element of the collection corresponds to a reference
'   Each element in the collection is a dictionary with 6 elements:
'   4.1 STR_CONFIGKEY_REFERENCE_NAME is a dictionary element with a string item, the name of the reference
'   4.2 STR_CONFIGKEY_REFERENCE_DESCRIPTION is a dictionary element with a string item, the description of the reference
'   4.3 STR_CONFIGKEY_REFERENCE_GUID is a dictionary element with a string item, the GUID of the reference
'   4.4 STR_CONFIGKEY_REFERENCE_MAJOR is a dictionary element with a long item, the major version of the reference
'   4.5 STR_CONFIGKEY_REFERENCE_MINOR is a dictionary element with a long item, the minor version number of the reference
'   4.6 STR_CONFIGKEY_REFERENCE_PATH is a dictionary element with a string item, the path to the reference
'
' An example printed by DisplayConfig:
'VBAProject Name        ExcelRainMan
'
'Base Path              C:\Users\Owner\OneDrive\Documents\Excel\ExcelRainMan\Test Project
'
'Module Paths
'ExcelRainManProject    ExcelRainManProject.bas
'UtilityRoutines        UtilityRoutines.bas
'SearchRoutines         SearchRoutines.bas
'FormatRoutines         FormatRoutines.bas
'ErrorRoutines          ErrorRoutines.bas
'ErrorLogClass          ErrorLogClass.cls
'TextFileClass          TextFileClass.cls
'VBARoutines            VBARoutines.bas
'TableArrayClass        TableArrayClass.cls
'PerformanceRoutines    PerformanceRoutines.bas
'PerformanceClass       PerformanceClass.cls
'
'References
'1            Name          Scripting
'1            Description   Microsoft Scripting Runtime
'1            GUID          {420B2830-E718-11CF-893D-00A0C9054228}
'1            Major 1
'1            Minor 0
'2            Name          VBIDE
'2            Description   Microsoft Visual Basic for Applications Extensibility 5.3
'2            GUID          {0002E157-0000-0000-C000-000000000046}
'2            Major 5
'2            Minor 3
'

Public Sub DisplayConfig()

    ' This routine is used for learning how the config data structure is built
    
    ' Version 1.0.4
    ' Refactored GetReferences out of MakeConfigFile
    
    Const RoutineName As String = Module_Name & "DisplayConfig"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    Dim ConfigEntry As Variant
    For Each ConfigEntry In pConfig
        Select Case ConfigEntry
        Case "VBAProject Name"
            Debug.Print ConfigEntry, pConfig(ConfigEntry)
        Case "Base Path"
            Debug.Print ConfigEntry, pConfig(ConfigEntry)
        Case "Module Paths"
            Debug.Print ConfigEntry
            Dim ModPath As Variant
            For Each ModPath In pConfig(ConfigEntry)
                Debug.Print ModPath, pConfig(ConfigEntry)(ModPath)
            Next ModPath
        Case "References"
            Debug.Print ConfigEntry
            Dim I As Long
            For I = 1 To pConfig(ConfigEntry).Count
                Dim RefElement As Variant
                For Each RefElement In pConfig(ConfigEntry)(I)
                    Debug.Print I, RefElement, pConfig(ConfigEntry)(I)(RefElement)
                Next RefElement
            Next I
        Case Else
            Debug.Print "Shouldn't get here"
        End Select
        Debug.Print
    Next ConfigEntry
    Stop
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    Stop
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' DisplayConfig

Private Sub Class_Initialize()
    Set pConfig = New Dictionary
    Set pProject = Nothing
End Sub                                          ' Class_Initialize

Public Property Get ConfigFile() As String
    '// The configuration file. Used for file paths

    ConfigFile = pConfigFile
End Property                                     ' Get ConfigFile

Public Property Let ConfigFile(ByVal NewFile As String)
    '// The configuration file. Used for file paths

    pConfigFile = NewFile
End Property                                     ' Let ConfigFile

Public Property Get Project() As VBProject
    '// The VBProject context for this configuration. Used for file paths

    Set Project = pProject
End Property                                     ' Get Project

Public Property Let Project(ByVal newProject As VBProject)
    '// The VBProject context for this configuration. Used for file paths

    Set pProject = newProject
    pConfig(STR_CONFIGKEY_PROJECTNAME) = newProject.Name
End Property                                     ' Let Project

Public Property Get VBAProjectName() As String
    '// Get the VBAProject name as determined by the configuration

    If pConfig.Exists(STR_CONFIGKEY_PROJECTNAME) Then
        VBAProjectName = pConfig(STR_CONFIGKEY_PROJECTNAME)
    Else
        VBAProjectName = vbNullString
    End If

End Property                                     ' Get VBAProjectName

Public Property Get VBAProjectNameDeclared() As Boolean
    '// Check if the VBAProject name is actually declared in the configuration
    '// This distinguishes between VBAProjectName = vbnullstring and not being set at all.
    
    VBAProjectNameDeclared = pConfig.Exists(STR_CONFIGKEY_PROJECTNAME)
End Property                                     ' Get VBAProjectNameDeclared

Public Sub VBAProjectNameRemove()
    If pConfig.Exists(STR_CONFIGKEY_PROJECTNAME) Then
        '// Remove the VBAProject name declaration from the configuration
        '// This makes VBAProjectNameDeclare = false
        
        pConfig.Remove STR_CONFIGKEY_PROJECTNAME
    End If
End Sub                                          ' VBAProjectNameRemove

Public Property Get BasePath() As String
    '// Get the base path as determined by the configuration

    If pConfig.Exists(STR_CONFIGKEY_BASEPATH) Then
        BasePath = pConfig(STR_CONFIGKEY_BASEPATH)
    Else
        BasePath = vbNullString
    End If
End Property                                     ' Get BasePath

Public Property Let BasePath(ByVal newBasePath As String)
    '// Set the configuration base path
    
    pConfig(STR_CONFIGKEY_BASEPATH) = newBasePath
End Property                                     ' Let BasePath

Public Property Get BasePathDeclared() As Boolean
    '// Check if the base path is actually declared in the configuration
    '// This distinguishes between BasePath = vbnullstring and not being set at all.
    
    BasePathDeclared = pConfig.Exists(STR_CONFIGKEY_BASEPATH)
End Property                                     ' Get BasePathDeclared

Public Sub BasePathRemove()
    '// Remove the base path declaration from the configuration
    '// This makes BasePathDeclared = false
    
    If pConfig.Exists(STR_CONFIGKEY_BASEPATH) Then
        pConfig.Remove STR_CONFIGKEY_BASEPATH
    End If
End Sub                                          ' BasePathRemove

Public Property Get ModuleNames() As Variant()
    '// List all the names of the modules which have been assigned a path in the
    '// configuration
    
    ModuleNames = ModulePathsDictReadOnly.Keys
End Property                                     ' Get ModuleNames

Public Property Get ModulePathDeclared(ByVal ModuleName As String) As Boolean
    '// Check if a path for the given module is declared in the configuration
    
    ModulePathDeclared = ModulePathsDictReadOnly.Exists(ModuleName)
End Property                                     ' Get ModulePathDeclared

Public Property Get ModulePath(ByVal ModuleName As String) As String
    '// Get the module path for a module as determined by the configuration
    
    ModulePath = ModulePathsDictReadOnly(ModuleName)
End Property                                     ' Get ModulePath

Public Property Let ModulePath(ByVal ModuleName As String, ByVal newPath As String)
    '// Set the configuration module path for a module
    
    ModulePathsDictForWriting(ModuleName) = newPath
End Property                                     ' Let ModulePath

Public Sub ModulePathRemove(ByVal ModuleName As String)
    '// Remove the module path declaration for a module from the configuration
    '// This makes ModuleDeclared(moduleName) = false
    
    ModulePathsDictReadOnly.Remove ModuleName
End Sub                                          ' ModulePathRemove

Private Property Get ModulePathsDictReadOnly() As Dictionary
    '// Get a dictionary entry that represents the module paths configuration
    '// The entry may not actually be the configuration dictionary so
    '// don't write to it.
    
    If pConfig.Exists(STR_CONFIGKEY_MODULEPATHS) Then
        Set ModulePathsDictReadOnly = pConfig(STR_CONFIGKEY_MODULEPATHS)
    Else
        Set ModulePathsDictReadOnly = New Dictionary
    End If

End Property                                     ' Get ModulePathsDictReadOnly

Public Property Get ModuleFullPath(ByVal ModuleName As String) As String
    '// Evaluate a module's absolute module path as determined by the configuration
    '// and the path of the project.

    ModuleFullPath = FSO.BuildPath(BasePath, ModulePath(ModuleName))
    
    If FSO.GetDriveName(ModuleFullPath) = vbNullString Then
        ModuleFullPath = FSO.BuildPath(FSO.GetParentFolderName(pConfigFile), ModuleFullPath)
    End If
    
    ModuleFullPath = FSO.GetAbsolutePathName(ModuleFullPath)

End Property                                     ' Get ModuleFullPath

Public Sub ModulePathsRemove()
    '// Remove the entire Module Paths declaration from the configuration
    
    If pConfig.Exists(STR_CONFIGKEY_MODULEPATHS) Then
        pConfig.Remove STR_CONFIGKEY_MODULEPATHS
    End If
End Sub                                          ' ModulePathsRemove

Public Property Get ModulePathsDictForWriting() As Dictionary
    '// Get the dictionary of the module paths configuration.
    '// If the module paths is not yet declared, this will create
    '// an empty declaration
    
    If Not pConfig.Exists(STR_CONFIGKEY_MODULEPATHS) Then
        Set pConfig(STR_CONFIGKEY_MODULEPATHS) = New Dictionary
    End If
    Set ModulePathsDictForWriting = pConfig(STR_CONFIGKEY_MODULEPATHS)

End Property                                     ' Get ModulePathsDictForWriting

Public Property Get ReferencesCount() As Long

    ' This routine returns the number of references in pConfig
    
    ' Version 1.0.4
    ' Refactored GetReferences out of MakeConfigFile
    
    Const RoutineName As String = Module_Name & "ReferencesCount"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    If pConfig.Exists(STR_CONFIGKEY_REFERENCES) Then
        ReferencesCount = pConfig(STR_CONFIGKEY_REFERENCES).Count
    Else
        ReferencesCount = 0
    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Property
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Property                                     ' ReferencesCount

Public Property Get ReferenceName(ByVal Index As Long) As String
    '// Get the name of a reference declared in the references configuration

    Dim collRefs    As Collection
    Dim dictRef     As Dictionary

    Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    Set dictRef = collRefs(Index)
    ReferenceName = dictRef(STR_CONFIGKEY_REFERENCE_NAME)
End Property                                     ' Get ReferenceName

Public Sub ReferencesUpdateFromVBRef(ByVal VBRef As Reference)
    '// Update the references list with a reference from the VBE.
    '// If the reference name is already in the list, that entry is overwritten
    '// otherwise a new entry is created.
    
    Dim lngIndex As Long
    For lngIndex = 1 To ReferencesCount
        If ReferenceName(lngIndex) = VBRef.Name Then
            ReferenceSetFromVBRef lngIndex, VBRef
            Exit Sub                             ' Assume there are no duplicate entries
        End If
    Next lngIndex

    ReferenceAddFromVBRef VBRef
End Sub                                          ' ReferencesUpdateFromVBRef

Private Sub ReferenceSetFromVBRef( _
        ByVal Index As Long, _
        ByVal VBRef As Reference)
    '// Set a reference in the references list to represent the given VB reference.
    
    ' Version 1.0.4
    ' Refactored GetReferences out of MakeConfigFile
    
    Const RoutineName As String = Module_Name & "ReferenceSetFromVBRef"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    Dim collRefs As Collection
    Dim dictRef As Dictionary
    
    Dim ErrorNumber As Long
    On Error Resume Next
    Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    ErrorNumber = Err.Number
    On Error GoTo ErrorHandler
                
    If ErrorNumber <> 0 Then
        Set collRefs = New Collection
        Set dictRef = New Dictionary
    End If
    
    Set dictRef = collRefs(Index)

    dictRef(STR_CONFIGKEY_REFERENCE_NAME) = VBRef.Name
    
    dictRef(STR_CONFIGKEY_REFERENCE_DESCRIPTION) = VBRef.Description
    
    If VBRef.Type = vbext_rk_TypeLib Then
        dictRef(STR_CONFIGKEY_REFERENCE_GUID) = VBRef.GUID
        dictRef(STR_CONFIGKEY_REFERENCE_MAJOR) = VBRef.Major
        dictRef(STR_CONFIGKEY_REFERENCE_MINOR) = VBRef.Minor
        
        If dictRef.Exists(STR_CONFIGKEY_REFERENCE_PATH) _
        Then dictRef.Remove STR_CONFIGKEY_REFERENCE_PATH
        
    Else
        dictRef(STR_CONFIGKEY_REFERENCE_PATH) = VBRef.FullPath
        
        If dictRef.Exists(STR_CONFIGKEY_REFERENCE_GUID) _
        Then dictRef.Remove STR_CONFIGKEY_REFERENCE_GUID
        
        If dictRef.Exists(STR_CONFIGKEY_REFERENCE_MAJOR) _
        Then dictRef.Remove STR_CONFIGKEY_REFERENCE_MAJOR
        
        If dictRef.Exists(STR_CONFIGKEY_REFERENCE_MINOR) _
        Then dictRef.Remove STR_CONFIGKEY_REFERENCE_MINOR
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ReferenceSetFromVBRef

Private Sub ReferenceAddFromVBRef(ByVal VBRef As Reference)
    '// Add a new reference to the reference list to represent the given reference.
    
    ' Version 1.0.4
    ' Refactored GetReferences out of MakeConfigFile
    
    Const RoutineName As String = Module_Name & "ReferenceAddFromVBRef"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    If Not pConfig.Exists(STR_CONFIGKEY_REFERENCES) Then
        Set pConfig(STR_CONFIGKEY_REFERENCES) = New Collection
    End If

    Dim collRefs As Collection
    Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    
    Dim dictRef As Dictionary
    Set dictRef = New Dictionary
    collRefs.Add dictRef

    ReferenceSetFromVBRef collRefs.Count, VBRef
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ReferenceAddFromVBRef

Public Sub ReferenceRemove(ByVal Index As Long)
    '// Remove a reference declaration from the references configuration collection
    
    pConfig(STR_CONFIGKEY_REFERENCES).Remove Index
End Sub                                          ' ReferenceRemove

Public Sub ReferencesRemove()
    '// Remove the entire References declaration from the configuration
    
    If pConfig.Exists(STR_CONFIGKEY_REFERENCES) Then
        pConfig.Remove STR_CONFIGKEY_REFERENCES
    End If
End Sub                                          ' ReferencesRemove

Public Sub ReferenceAddToVBRefs( _
       ByVal Index As Long, _
       ByVal VBRefs As References)
    '// Add a reference in the references configuration list to the VBE references
    '// list. Will delete any existing reference in the VBE references list which
    '// has the same name.

    Dim collRefs As Collection
    Set collRefs = pConfig(STR_CONFIGKEY_REFERENCES)
    
    Dim dictRef As Dictionary
    Set dictRef = collRefs(Index)

    Dim boolLibTypeValid As Boolean
    boolLibTypeValid = _
                     dictRef.Exists(STR_CONFIGKEY_REFERENCE_GUID) And _
                     dictRef.Exists(STR_CONFIGKEY_REFERENCE_MAJOR) And _
                     dictRef.Exists(STR_CONFIGKEY_REFERENCE_MINOR)
        
    Dim boolPathTypeValid As Boolean
    boolPathTypeValid = dictRef.Exists(STR_CONFIGKEY_REFERENCE_PATH)

    If Not (boolLibTypeValid Or boolPathTypeValid) Then
        ' TODO This reference isn't valid. Should at least warn the user?
        Exit Sub
    End If

    If CheckNameInCollection(ReferenceName(Index), VBRefs) Then
        VBRefs.Remove VBRefs(ReferenceName(Index))
    End If

    If boolLibTypeValid Then
        VBRefs.AddFromGuid _
        GUID:=dictRef(STR_CONFIGKEY_REFERENCE_GUID), _
        Major:=dictRef(STR_CONFIGKEY_REFERENCE_MAJOR), _
        Minor:=dictRef(STR_CONFIGKEY_REFERENCE_MINOR)
    Else                                         'PathTypeValid
        VBRefs.AddFromFile _
        FileName:=dictRef(STR_CONFIGKEY_REFERENCE_PATH)
    End If

End Sub                                          ' ReferenceAddToVBRefs

Public Function ReferenceExists(ByVal Ref As Reference) As Boolean

    ' This routine returns True if Ref exists in pConfig
    
    ' Version 1.0.4
    ' Refactored GetReferences out of MakeConfigFile
    
    Const RoutineName As String = Module_Name & "ReferenceExists"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    ReferenceExists = False
    
    If pConfig.Exists(STR_CONFIGKEY_REFERENCES) Then
        If pConfig(STR_CONFIGKEY_REFERENCES).Count = 0 Then
            pConfig.Remove (STR_CONFIGKEY_REFERENCES)
            ReferenceExists = False
            Exit Function
        End If
    Else
        ReferenceExists = False
        Exit Function
    End If
    
    Dim I As Long
    For I = 1 To pConfig(STR_CONFIGKEY_REFERENCES).Count
        If pConfig(STR_CONFIGKEY_REFERENCES)(I)(STR_CONFIGKEY_REFERENCE_GUID) = Ref.GUID Then
            ReferenceExists = True
            Exit Function
        End If
    Next I

    '    Dim ErrorNumber As Long
    '    Dim RefCount As Long
    '    On Error Resume Next
    '    RefCount = pConfig(STR_CONFIGKEY_REFERENCES).Count
    '    ErrorNumber = Err.Number
    '    On Error GoTo ErrorHandler
    '
    '    If ErrorNumber <> 0 Then
    '        ReferenceExists = False
    '    End If
    '
    '    If RefCount = 0 Then
    '        pConfig(STR_CONFIGKEY_REFERENCES).RemoveAll
    '    Else
    '        Dim I As Long
    '        For I = 1 To RefCount
    '            If pConfig(STR_CONFIGKEY_REFERENCES)(I)(STR_CONFIGKEY_REFERENCE_GUID) = Ref.GUID Then
    '                ReferenceExists = True
    '                Exit Function
    '            End If
    '        Next I
    '    End If

    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ReferenceExists

Public Sub ReferencesAddToVBRefs(ByVal VBRefs As References)
    '// Add all references in the references configuration list to the VBE
    '// references list. It could be argued that this doesn't logically belong
    '// in this class, but it's here because it's fairly trivial addition to
    '// ReferenceAddToVBRefs.
    
    Dim lngIndex As Long
    For lngIndex = 1 To ReferencesCount
        ReferenceAddToVBRefs lngIndex, VBRefs
    Next lngIndex
End Sub                                          ' ReferencesAddToVBRefs

Public Sub ReadFromProjectConfigFile()
    '// Reads and deserializes configuration from the configuration file for the
    '// current project.
    
    '// Read JSON file and build the configuration dictionary
    If FSO.FileExists(pConfigFile) Then
        Dim tsConfigStream As Scripting.TextStream
        Set tsConfigStream = FSO.OpenTextFile(pConfigFile, ForReading)
        
        Dim strConfigJson As String
        strConfigJson = tsConfigStream.ReadAll()
        
        tsConfigStream.Close
        Set pConfig = JsonConverter.ParseJson(strConfigJson)
    Else
        Set pConfig = New Dictionary
    End If

End Sub                                          ' ReadFromProjectConfigFile

Public Sub WriteToProjectConfigFile()
    '// Serializes and writes configuration to the configuration file for the
    '// current project.
    
    '// Write JSON to file
    Dim strConfigJson As String
    strConfigJson = JsonConverter.ConvertToJson(pConfig, vbTab)
    strConfigJson = strConfigJson & vbNewLine
    
    Dim tsConfigStream As Scripting.TextStream
    Set tsConfigStream = FSO.CreateTextFile(ConfigFile, True)
    
    tsConfigStream.Write strConfigJson
    '    tsConfigStream.Close

End Sub                                          ' WriteToProjectConfigFile


