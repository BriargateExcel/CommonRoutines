Attribute VB_Name = "CodeImportExport"
'
'
'
' From: https://github.com/spences10/VBA-IDE-Code-Export
'
'
'
Option Explicit

Const Module_Name As String = "CodeImportExport."

'// Updates the configuration file for the current active project.
'// * Entries for modules not yet declared in the configuration file as created.
'// * Modules listed in the configuration file which are not found are prompted
'//   to be deleted from the configuration file.
'// * The current loaded references are used to update the configuration file.
'// * References in the configuration file whic hare not loaded are prompted to
'//   be deleted from the configuration file.

Private pFormCanceled As Boolean
Private pFormDeleted As Boolean
Private pFormExportFlag As Boolean

' Version 1.0
' Refactored the config file and base path finders into separate routines
' Added document properties for the locations of the configuration file and the base path

Private Const ConfigFileDocProp As String = "ConfigFile"
Private Const BasePathDocProp As String = "BasePath"

Public FSO As FileSystemObject

Public Sub MakeConfigFile()

    ' This routine builds the configuration file describing the export and import contents
    
    ' Version 1.0
    ' Refactored GetConfigFile and GetBasePath out of MakeConfigFile
    ' Version 1.0.3
    ' Refactored GetModules out of MakeConfigFile
    ' Version 1.0.4
    ' Refactored GetReferences out of MakeConfigFile
    
    Const RoutineName As String = Module_Name & "MakeConfigFile"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        ResetPerformance
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    pFormExportFlag = False
    
    ' Determine for which VBA Project to build config file
    Dim prjActProj As VBProject
    Set prjActProj = GetProject("Build Config File")
    If prjActProj Is Nothing Then GoTo exitSub
    
    Dim Wkbk As Workbook
    Set FSO = New FileSystemObject
    Set Wkbk = Workbooks(FSO.GetFileName(prjActProj.FileName))
        
    Dim config As clsConfiguration
    Set config = New clsConfiguration
    
    ' Get the config file and its contents (if it has contents)
    Dim ConfigFile As String
    ConfigFile = GetConfigFile(Wkbk, prjActProj) ' Version 1.0
    If ConfigFile = "No configuration file directory selected" Then
        MsgBox "No configuration file selected and no directory " & _
               "for a new configuration file selected." & _
               vbCrLf & _
               "No configuration file created.", _
               vbOKOnly Or vbCritical, _
               "No Configuration File or Directory Selected"
        Exit Sub
    Else
        config.ConfigFile = ConfigFile
        config.ReadFromProjectConfigFile
    End If
    config.Project = prjActProj
    ' Now have config file location and contents if it has contents
    
    ' Get the base path
    Dim BasePath As String
    BasePath = GetBasePath(Wkbk)                 ' Version 1.0
    If BasePath = "No base path selected" Then
        MsgBox "No base path selected. No configuration file created.", _
               vbOKOnly Or vbCritical, _
               "No Base Path Selected"
        Exit Sub
    Else
        config.BasePath = BasePath
    End If
    ' Now have the base path defined
    
    '// Generate entries for modules not yet listed
    GetModules prjActProj, config

    '// Generate entries for references in the current VBProject
    GetReferences prjActProj, config
    
    '// Write changes to config file
    If FSO.FileExists(config.ConfigFile) Then
        Dim Response As Long
        Response = MsgBox("The configuration file already exists. " & _
                          vbCrLf & _
                          config.ConfigFile & _
                          "Would you like to overwrite it?", _
                          vbYesNo Or vbExclamation, _
                          "Configuration File Already Exists")
        Select Case Response
        Case vbYes
            UpdateConfigFile config
        Case vbNo
            MsgBox "The configuration file was not overwritten", _
                   vbOKOnly, _
                   "Configuration File Not Overwritten"
        End Select

    Else
        UpdateConfigFile config
    End If
    ' Config file has been updated

exitSub:
    Exit Sub
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName
End Sub                                          ' MakeConfigFile

Public Sub Export()

    '// Exports code modules and cleans the current active VBProject as specified
    '// by the project's configuration file.
    '// * Any code module in the VBProject which is listed in the configuration
    '//   file is exported to the configured path.
    '// * code modules which were exported are deleted or cleared.
    '// * References loaded in the Project which are listed in the configuration
    '//   file is deleted.
    
    ' Version 1.0
    ' Modified Export to use GetConfigFile
    
    Const RoutineName As String = Module_Name & "Export"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        ResetPerformance                         ' Only in the Main routine; delete in all others
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    On Error GoTo ErrorHandler

    pFormExportFlag = True

    Dim prjActProj As VBProject
    Set prjActProj = GetProject("Export")
    If prjActProj Is Nothing Then GoTo NoAction

    Dim config As clsConfiguration
    Set config = New clsConfiguration
    config.Project = prjActProj
    
    Set FSO = New FileSystemObject
    Dim Wkbk As Workbook
    Set Wkbk = Workbooks(FSO.GetFileName(prjActProj.FileName))

    Dim SelectedFile As String
    SelectedFile = GetConfigFile(Wkbk, prjActProj, False) ' Version 1.0
    If SelectedFile = "No configuration file directory selected" Then
        MsgBox "No configuration file selected." & _
               vbCrLf & _
               "Export failed.", _
               vbOKOnly Or vbCritical, "No Configuration File Selected"
        Exit Sub
    Else
        config.ConfigFile = SelectedFile
    End If
    
    config.ReadFromProjectConfigFile

    '// Export all modules listed in the configuration
    Dim varModuleName As Variant
    For Each varModuleName In config.ModuleNames
        Dim strModuleName As String
        strModuleName = varModuleName
        ' TODO Provide a warning if module listed in configuration is not found
        If CheckNameInCollection(strModuleName, prjActProj.VBComponents) Then
            Dim comModule As VBComponent
            Set comModule = prjActProj.VBComponents(strModuleName)
            
            EnsurePath config.ModuleFullPath(strModuleName)
            Dim Dest As String
            Dest = config.ModuleFullPath(strModuleName)
            comModule.Export Dest

            If pFormDeleted Then
                If comModule.Type = vbext_ct_Document Then
                    comModule.CodeModule.DeleteLines 1, comModule.CodeModule.CountOfLines
                Else
                    prjActProj.VBComponents.Remove comModule
                End If
            End If
        End If
    Next varModuleName

    '// Remove all references listed
    If pFormDeleted Then
        Dim lngIndex As Long
        For lngIndex = 1 To config.ReferencesCount
            If CheckNameInCollection(config.ReferenceName(lngIndex), prjActProj.References) Then
                prjActProj.References.Remove prjActProj.References(config.ReferenceName(lngIndex))
            End If
        Next lngIndex
    End If

    MsgBox "All modules successfully exported", _
           vbOKOnly Or vbInformation, _
           "Modules Exported Successfully"
NoAction:
    Exit Sub
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName
End Sub                                          ' Export

Public Sub Import()

    '// Imports textual data from the file system such as VBA code to build the
    '// current active VBProject as specified in its configuration file.
    '// * Each code module file listed in the configuration file is imported into
    '//   the VBProject.
    '//   TODO Warn the user of overwrites
    '//   Modules with the same name are overwritten.
    '// * All references declared in the configuration file are
    '//   loaded into the project.
    '// * The project name is set to the project name specified by
    '//   the configuration file.
    ' Version 1.0
    ' Modified Import to use GetConfigFile
    
    Const RoutineName As String = Module_Name & "Import"
    On Error GoTo ErrorHandler
    
    pFormExportFlag = False

    Dim prjActProj As VBProject
    Set prjActProj = GetProject("Import")
    If prjActProj Is Nothing Then Exit Sub
    
    If prjActProj.Name = "Personal" Then
        MsgBox "Can not import PERSONAL.xlsb because " & _
               "the import code is in PERSONAL.xlsb", _
               vbOKOnly Or vbCritical, _
               "Can Not Import PERSONAL.xlsb"
        Exit Sub
    End If

    Dim config As clsConfiguration
    Set config = New clsConfiguration
    config.Project = prjActProj
    
    Dim Ary As Variant
    Ary = Split(prjActProj.FileName, "/")
    
    Dim FileName As String
    FileName = Ary(UBound(Ary, 1))
    
    Dim Wkbk As Workbook
    Dim ErrorNumber As Long
    On Error Resume Next
    Set Wkbk = Workbooks(FileName)
    ErrorNumber = Err.Number
    On Error GoTo ErrorHandler
    If ErrorNumber <> 0 Then
        MsgBox "You must save the file before you can import into it", _
               vbOKOnly Or vbCritical, _
               "File Not Saved"
        Exit Sub
    End If
    
    Dim SelectedFile As String
    SelectedFile = GetConfigFile(Wkbk, prjActProj, False) ' Version 1.0
    If SelectedFile = "No configuration file directory selected" Then
        MsgBox "No configuration file selected." & _
               vbCrLf & _
               "Import failed.", _
               vbOKOnly Or vbCritical, "No Configuration File Selected"
        Exit Sub
    Else
        config.ConfigFile = SelectedFile
    End If
    
    config.ReadFromProjectConfigFile

    '// Import code from listed module files
    Dim varModuleName As Variant
    For Each varModuleName In config.ModuleNames
        Dim strModuleName As String
        strModuleName = varModuleName
        ImportModule prjActProj, strModuleName, config.ModuleFullPath(strModuleName)
    Next varModuleName

    '// Add references listed in the config file
    config.ReferencesAddToVBRefs prjActProj.References
    
    '// Set the VBA Project name
    If config.VBAProjectNameDeclared Then
        prjActProj.Name = config.VBAProjectName
    End If

    MsgBox "All modules successfully imported", _
           vbOKOnly Or vbInformation, _
           "Modules Imported Successfully"
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    DisplayError RoutineName
End Sub                                          ' Import

Private Function GetProject(ByVal TitleText As String) As VBProject
    ' Version 1.0
    ' Refactored this out of MakeConfigFile
    
    Const RoutineName As String = Module_Name & "GetProject"
    On Error GoTo ErrorHandler
    
    GitForm.ProjectList.Clear
    
    Dim VBAProj As Variant
    For Each VBAProj In Application.VBE.VBProjects
        GitForm.ProjectList.AddItem VBAProj.Name
    Next VBAProj
    
    If GitForm.ProjectList.ListCount > 0 And _
       GitForm.ProjectList.List(0) = "Personal" _
       Then
        GitForm.ProjectList.Text = GitForm.ProjectList.List(1)
    Else
        GitForm.ProjectList.Text = GitForm.ProjectList.List(0)
    End If
    GitForm.Caption = "Select the VBA Project to " & TitleText
    GitForm.Show
    If pFormCanceled Then
        ' Either Cancel button or dialog close button (red X) selected
        MsgBox TitleText & " canceled by user", _
               vbOKOnly Or vbInformation, _
               "Cancel Selected"
        Set GetProject = Nothing
        Exit Function
    Else
        Dim prjActProj As VBProject
        Set prjActProj = Application.VBE.VBProjects(GitForm.ProjectList.Value)
    End If

    If prjActProj.Protection = 1 Then
        MsgBox "This project is protected, not possible to export the code"
        Exit Function
    End If
    
    Set GetProject = prjActProj
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetProject

Private Function GetConfigFile( _
        ByVal Wkbk As Workbook, _
        ByVal prjActProj As VBProject, _
        Optional ByVal SelectDirectory As Boolean = True _
        ) As String

    ' This routine checks the document properties to see if the JSON file has already been stored
    ' It starts from that location when asking the user for the JSON file
    ' It stores the JSON file location back into the document properties
    ' It raises an error if the user does not select a configuration file
    '   or a directory to put a new configuration file
    ' Version 1.0
    ' Refactored this out of MakeConfigFile
    ' Version 1.0.2
    ' Added conditional compilation for document properites
    
    Const RoutineName As String = Module_Name & "GetConfigFile"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    ' Get the JSON file location from the document properties
    Dim InitialConfigFile As String
    InitialConfigFile = GetDocPropConfigFile(Wkbk)
    
    ' Confirm the JSON file with the user
    Dim SelectedFile As String
    SelectedFile = GetUserConfigFile(InitialConfigFile)
    If SelectedFile <> "No file selected" Then
        #If DocProp = 1 Then
            If Not SetProperty(ConfigFileDocProp, PropertyLocationCustom, SelectedFile, False, Wkbk) Then
                ' This is not fatal
                MsgBox "Unable to store document property", _
                       vbOKOnly Or vbCritical, _
                       "Unable to Store Document Property"
            End If
        #End If
    Else
        If Not SelectDirectory Then
            ' This is fatal
            GetConfigFile = "No configuration file directory selected"
            Exit Function
        End If
        ' The user didn't select an existing configuration file
        ' Give the user an opportunity to select a directory for a new configuration file
        MsgBox "No configuration file selected. " & _
               vbCrLf & _
               "Select directory for new configuration file.", _
               vbOKOnly Or vbCritical, _
               "No Configuration File"
               
        Dim Path As String
        Path = FSO.GetParentFolderName(Wkbk.Path)
        
        Dim SelectedConfigDirectory As String
        SelectedConfigDirectory = GetConfigDirectory(Path)
        If SelectedConfigDirectory = "No directory selected" Then
            ' This is fatal
            GetConfigFile = "No configuration file directory selected"
            Exit Function
        Else
            SelectedFile = _
                         SelectedConfigDirectory & _
                         Application.PathSeparator & _
                         prjActProj.Name & ".MakeFile.json"
        End If
    
        #If DocProp = 1 Then
            Dim GoodProperty As Boolean
            Dim ErrorNumber As Long
            On Error Resume Next
            GoodProperty = SetProperty(ConfigFileDocProp, PropertyLocationCustom, SelectedFile)
            ErrorNumber = Err.Number
            On Error GoTo ErrorHandler
    
            If ErrorNumber = 0 Then
                If Not GoodProperty Then
                    ' This is not fatal
                    MsgBox "Unable to store configuration file path in " & _
                           vbCrLf & _
                           Wkbk.Name & "'s document properties", _
                           vbOKOnly Or vbCritical, _
                           "Unable to Store Config File Path"
                End If
            End If
        #End If
    End If
    
    GetConfigFile = SelectedFile
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetConfigFile

Private Sub GetModules( _
        ByVal prjActProj As VBProject, _
        ByRef config As clsConfiguration)

    ' This routine gathers a list of the modules in this project
    ' Compares that list with the existing config file
    ' Modifies the list of modules if the user desires
    
    ' Version 1.0.3
    ' Refactored GetModules out of MakeConfigFile
    
    Const RoutineName As String = Module_Name & "GetModules"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    Dim collAddList As Collection
    Set collAddList = New Collection
    
    Dim strAddListStr As String
    strAddListStr = vbNullString
    
    Dim strModuleName As String
    Dim comModule As VBComponent
    For Each comModule In prjActProj.VBComponents
        Dim boolCreateNewEntry  As Boolean
        boolCreateNewEntry = _
                           ExportableModule(comModule) And _
                           Not config.ModulePathDeclared(comModule.Name)

        If boolCreateNewEntry Then
            config.ModulePath(comModule.Name) = comModule.Name & "." & FileExtension(comModule)
            collAddList.Add comModule.Name
            strAddListStr = strAddListStr & comModule.Name & vbNewLine
        End If
    Next comModule

    ' Ask the user if they want to add new modules to the config file
    If collAddList.Count > 0 Then
        Dim intUserResponse As Long
        intUserResponse = MsgBox( _
                          Prompt:= _
                          "There are some modules not listed in the configuration file which " & _
                          "exist in the current project. Would you like to " & _
                          "add these modules to the configuration file?" & _
                          vbNewLine & _
                          "Note: All modules are listed if there is no existing configuration file" & _
                          vbNewLine & _
                          "New modules:" & vbNewLine & _
                          strAddListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="New Modules")

        If intUserResponse = vbYes Then
            Dim varModuleName As Variant
            For Each varModuleName In collAddList
                strModuleName = varModuleName
                config.ModulePath(strModuleName) = strModuleName & "." & FileExtension(prjActProj.VBComponents(strModuleName))
            Next varModuleName
        End If
    End If
    
    '// Ask user if they want to delete entries for missing modules
    ' Create the list of modules to potentially delete
    Dim collDeleteList As Collection
    Set collDeleteList = New Collection

    Dim strDeleteListStr As String
    strDeleteListStr = vbNullString

    For Each varModuleName In config.ModuleNames
        strModuleName = varModuleName

        Dim boolDeleteModule As Boolean
        boolDeleteModule = True

        If CheckNameInCollection(strModuleName, prjActProj.VBComponents) Then
            If ExportableModule(prjActProj.VBComponents(strModuleName)) Then
                boolDeleteModule = False
            End If
        End If

        If boolDeleteModule Then
            collDeleteList.Add strModuleName
            strDeleteListStr = strDeleteListStr & strModuleName & vbNewLine
        End If
    Next varModuleName
    ' Now have a list of modules to potentially delete

    ' Ask the user if they want to delete any modules
    If collDeleteList.Count > 0 Then
        intUserResponse = MsgBox( _
                          Prompt:= _
                          "There are some modules listed in the configuration file which " & _
                          "haven't been found in the current project. Would you like to " & _
                          "remove these modules from the configuration file?" & _
                          vbNewLine & _
                          vbNewLine & _
                          "Missing modules:" & vbNewLine & _
                          strDeleteListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="Missing Modules")

        If intUserResponse = vbYes Then
            For Each varModuleName In collDeleteList
                strModuleName = varModuleName
                config.ModulePathRemove strModuleName
            Next varModuleName
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' GetModules

Private Function GetDocPropConfigFile(ByVal Wkbk As Workbook) As String

    ' Version 1.0
    ' Part of refactoring GetConfigFile
    ' Version 1.0.1
    ' Put this routine into normal form
    ' Version 1.0.2
    ' Added conditional compilation for document properites
    
    Const RoutineName As String = Module_Name & "GetDocPropConfigFile"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    #If DocProp = 1 Then
        Dim ErrorNumber As Long
        Dim GoodName As Boolean
        
        On Error Resume Next
        GoodName = NameExists(ConfigFileDocProp, Wkbk)
        ErrorNumber = Err.Number
        On Error GoTo 0
    
        If ErrorNumber = 0 Then
            If GoodName Then
                GetDocPropConfigFile = GetProperty(ConfigFileDocProp, PropertyLocationCustom, Wkbk)
            Else
                GetDocPropConfigFile = vbNullString
            End If
        Else
            GetDocPropConfigFile = vbNullString
        End If
    #End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetDocPropConfigFile

Private Function GetUserConfigFile(ByVal InitialFile As String) As String
    
    ' Open the file dialog and capture the folder's path
    ' Version 1.0
    ' Part of refactoring GetConfigFile
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "GetUserConfigFile"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "json", "*.json"
        .AllowMultiSelect = False
        .InitialFileName = InitialFile
        .Title = "Choose Configuration File"
            
        Dim Response As Long
        Response = .Show
        If Response <> 0 Then
            GetUserConfigFile = .SelectedItems(1)
        Else
            GetUserConfigFile = "No file selected"
        End If
            
    End With
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetUserConfigFile

Private Function GetConfigDirectory(ByVal InitialDirectory As String) As String
    
    ' Open the file dialog and capture the folder's path
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "GetConfigDirectory"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Choose Configuration File Directory"
        .InitialFileName = InitialDirectory
        Dim Response As Long
        Response = .Show
        If Response <> 0 Then
            GetConfigDirectory = .SelectedItems(1)
        Else
            GetConfigDirectory = "No directory selected"
        End If
            
    End With
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetConfigDirectory

Private Function GetBasePath( _
        ByVal Wkbk As Workbook _
        ) As String

    ' This routine checks the document properties to see if the base path has already been recorded
    ' It starts from that location when asking the user for the base path
    ' It stores the base path location back into the document properties
    
    ' Version 1.0
    ' Refactored out of MakeConfigFile then patterned  after GetConfigFile
    ' Version 1.0.2
    ' Added conditional compilation for document properites
        
    Const RoutineName As String = Module_Name & "GetBasePath"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    ' Get the base path location from the document properties
    Dim InitialBasePath As String
    InitialBasePath = GetDocPropBasePath(Wkbk)
    
    ' If the base path is stored, it is one folder level too deep
    '   for the directory selector
    ' Eliminate the last folder level
    InitialBasePath = FSO.GetParentFolderName(InitialBasePath)
    
    ' Confirm the base path with the user
    ' Or let the user select a new base path
    Dim SelectedBasePath As String
    SelectedBasePath = GetUserBasePath(InitialBasePath)
    If SelectedBasePath <> "No base path selected" Then
        #If DocProp = 1 Then
            If Not SetProperty(BasePathDocProp, PropertyLocationCustom, SelectedBasePath, False, Wkbk) Then
                ' This is not fatal
                MsgBox "Unable to set base path document property", _
                       vbOKOnly Or vbCritical, _
                       "Unable to Set Base Path Document Property"
            End If
        #End If
    Else
        ' User did not select a base path
        ' This is fatal
        GetBasePath = "No base path selected"
        Exit Function
    End If
    
    GetBasePath = SelectedBasePath
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetBasePath

Private Function GetDocPropBasePath(ByVal Wkbk As Workbook) As String
        
    ' Version 1.0
    ' Part of refactoring GetBasePath
    ' Version 1.0.1
    ' Put this routine into normal form
    ' Version 1.0.2
    ' Added conditional compilation for document properites
    
    Const RoutineName As String = Module_Name & "FunctionTemplate"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    #If DocProp = 1 Then
        Dim ErrorNumber As Long
        Dim GoodName As Boolean
        
        On Error Resume Next
        GoodName = NameExists(BasePathDocProp, Wkbk)
        ErrorNumber = Err.Number
        On Error GoTo 0
    
        If ErrorNumber = 0 Then
            If GoodName Then
                GetDocPropBasePath = GetProperty(BasePathDocProp, PropertyLocationCustom, Wkbk)
            Else
                GetDocPropBasePath = vbNullString
            End If
        Else
            GetDocPropBasePath = vbNullString
        End If
    #End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' FunctionTemplate

Private Function GetUserBasePath(ByVal InitialDirectory As String) As String
    
    ' Open the file dialog and capture the folder's path
    ' Version 1.0
    ' Part of refactoring GetBasePath
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "GetUserBasePath"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Choose Base Path Directory"
        .InitialFileName = InitialDirectory
        Dim Response As Long
        Response = .Show
        If Response <> 0 Then
            GetUserBasePath = .SelectedItems(1)
        Else
            GetUserBasePath = "No base path selected"
        End If
            
    End With
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' GetUserBasePath

Private Sub GetReferences( _
        ByVal prjActProj As VBProject, _
        ByRef config As clsConfiguration)

    ' This routine gathers a list of the references in this project
    ' Compares that list with the existing config file
    ' Modifies the list of reference if the user desires
    
    ' Version 1.0.4
    ' Refactored GetReferences out of MakeConfigFile
    
    Const RoutineName As String = Module_Name & "GetReferences"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    Dim collAddList As Collection
    Set collAddList = New Collection
    
    Dim strAddListStr As String
    strAddListStr = vbNullString
    
    Dim refReference As Reference
    For Each refReference In prjActProj.References
        If Not refReference.BuiltIn Then
            If ReferenceToAdd(refReference) Then
                If Not config.ReferenceExists(refReference) Then
                    collAddList.Add refReference
                    strAddListStr = strAddListStr & refReference.Name & vbNewLine
                End If
            End If
        End If
    Next refReference

    ' Ask the user if they want to add new references to the config file
    If collAddList.Count > 0 Then
        Dim intUserResponse As Long
        intUserResponse = MsgBox( _
                          Prompt:= _
                          "There are some references not listed in the configuration file which " & _
                          "exist in the current project. Would you like to " & _
                          "add these references to the configuration file?" & _
                          vbNewLine & _
                          "Note: if the configuration file doesn't already exist, this will be a list of all references" & _
                          vbNewLine & _
                          "New references:" & vbNewLine & _
                          strAddListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="New References")

        If intUserResponse = vbYes Then
            Dim I As Long
            I = 1
            Dim Ref As Variant
            For Each Ref In collAddList
                config.ReferencesUpdateFromVBRef Ref
                I = I + 1
            Next Ref
        End If
    End If
    
    '// Ask user if they want to delete entries for missing references
    ' Create the list of modules to potentially delete
    Dim collDeleteList As Collection
    Set collDeleteList = New Collection

    Dim strDeleteListStr As String
    strDeleteListStr = vbNullString
    
    Dim lngIndex As Long
    For lngIndex = config.ReferencesCount To 1 Step -1
        If Not CheckNameInCollection(config.ReferenceName(lngIndex), prjActProj.References) Then
            collDeleteList.Add lngIndex
            strDeleteListStr = strDeleteListStr & config.ReferenceName(lngIndex) & vbNewLine
        End If
    Next

    ' Ask the user if they want to delete any references
    If collDeleteList.Count > 0 Then
        intUserResponse = MsgBox( _
                          Prompt:="There are some references listed in the configuration file which " & _
                                   "haven't been found in the current project. Would you like to " & _
                                   "remove these references from the configuration file?" & vbNewLine & _
                                   vbNewLine & _
                                   "Missing references:" & vbNewLine & _
                                   strDeleteListStr, _
                          Buttons:=vbYesNo + vbDefaultButton2, _
                          Title:="Missing References")

        If intUserResponse = vbYes Then
            Dim varIndex As Variant
            For Each varIndex In collDeleteList
                lngIndex = varIndex
                config.ReferenceRemove lngIndex
            Next varIndex
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' GetReferences

Private Function ReferenceToAdd(ByVal ThisRef As Reference) As Boolean

    ' This routine determies if this reference needs to be saved
    ' Returns True if the reference must be saved
    ' Returns False if this reference does not need to be saved
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ReferenceToAdd"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    Dim NeededRef As Boolean
    Dim ErrorNumber As Long
    
    On Error Resume Next
    NeededRef = (ThisRef.Name = "stdole")
    ErrorNumber = Err.Number
    On Error GoTo 0
    If ErrorNumber <> 0 Then
        ReferenceToAdd = False                   ' Bad reference; skip
        Exit Function
    End If
    
    Select Case ThisRef.Name
    Case "stdole"
        ReferenceToAdd = False
    Case "MSForms"
        ReferenceToAdd = False
    Case "Office"
        ReferenceToAdd = False
    Case Else
        ReferenceToAdd = True
    End Select
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ReferenceToAdd

Private Sub ImportModule( _
        ByVal Project As VBProject, _
        ByVal ModuleName As String, _
        ByVal ModulePath As String)
    '// Import a VBA code module... how hard could it be right?
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ImportModule"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    Dim ErrorNumber As Long
    Dim comNewImport As VBComponent
    
    Dim NameToCheck As String
    On Error Resume Next
    NameToCheck = Project.VBComponents.Item(ModuleName).Name
    ErrorNumber = Err.Number
    On Error GoTo 0
    If ErrorNumber = 0 Then
        Select Case MsgBox( _
               "Module " & ModuleName & _
               " already exists. Overwrite it?", _
               vbYesNo Or vbExclamation, _
               "Module Already Exists")
        Case vbYes
            Dim VBC As VBComponent
            Set VBC = Project.VBComponents.Item(ModuleName)
            If Project.VBComponents.Item(ModuleName).Type <> vbext_ct_Document Then
                ' Can't remove a worksheet
                Project.VBComponents.Remove VBC
            End If
        Case vbNo
            Exit Sub
        End Select
    End If
    
    On Error Resume Next
    Set comNewImport = Project.VBComponents.Import(ModulePath)
    ErrorNumber = Err.Number
    If ErrorNumber = 60061 Then Exit Sub         ' Module already in use
    On Error GoTo 0
    
    If comNewImport.Name <> ModuleName Then
        If CheckNameInCollection(ModuleName, Project.VBComponents) Then

            Dim comExistingComp As VBComponent
            Set comExistingComp = Project.VBComponents(ModuleName)
            If comExistingComp.Type = vbext_ct_Document Then

                Dim modCodeCopy As CodeModule
                Set modCodeCopy = comNewImport.CodeModule
                
                Dim modCodePaste As CodeModule
                Set modCodePaste = comExistingComp.CodeModule
                
                modCodePaste.DeleteLines 1, modCodePaste.CountOfLines
                
                If modCodeCopy.CountOfLines > 0 Then
                    modCodePaste.AddFromString modCodeCopy.lines(1, modCodeCopy.CountOfLines)
                End If
                
                Project.VBComponents.Remove comNewImport

            Else
                comExistingComp.Name = comExistingComp.Name & "_remove"
                Project.VBComponents.Remove comExistingComp
                comNewImport.Name = ModuleName   ' TODO fails on work computer
                Project.VBComponents.Remove comExistingComp
            End If
        Else

            comNewImport.Name = ModuleName

        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' ImportModule

Private Function ExportableModule(ByVal comModule As VBComponent) As Boolean
    '// Is the given module exportable by this tool?
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ExportableModule"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    ExportableModule = _
                     (Not ModuleEmpty(comModule)) _
                     And _
                     (FileExtension(comModule) <> vbNullString)
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ExportableModule

Private Function ModuleEmpty(ByVal comModule As VBComponent) As Boolean
    '// Check if a code module is effectively empty.
    '// effectively empty should be functionally and semantically equivalent to
    '// actually empty.
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "ModuleEmpty"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    ModuleEmpty = True

    Dim lngNumLines As Long
    lngNumLines = comModule.CodeModule.CountOfLines
    
    Dim lngCurLine As Long
    For lngCurLine = 1 To lngNumLines
        Dim strCurLine As String
        strCurLine = comModule.CodeModule.lines(lngCurLine, 1)
        
        If Not (strCurLine = "Option Explicit" Or strCurLine = vbNullString) Then
            ModuleEmpty = False
            Exit Function
        End If
    Next lngCurLine
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' ModuleEmpty

Private Function FileExtension(ByVal comModule As VBComponent) As String
    '// The appropriate file extension for exporting the given module
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "FileExtension"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    Select Case comModule.Type
    Case vbext_ct_StdModule
        FileExtension = "bas"
    Case vbext_ct_ClassModule, vbext_ct_Document
        FileExtension = "cls"
    Case vbext_ct_MSForm
        FileExtension = "frm"
    Case Else
        FileExtension = vbNullString
    End Select
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Function
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Function                                     ' FileExtension

Private Sub EnsurePath(ByVal Path As String)
    '// Ensure path to a file exists. Creates missing folders.
    ' Version 1.0
    ' Added error handling
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "EnsurePath"
    On Error GoTo ErrorHandler
    
    Set FSO = New Scripting.FileSystemObject
    
    Dim strParentPath As String
    strParentPath = FSO.GetParentFolderName(Path)

    If strParentPath <> vbNullString Then
        EnsurePath strParentPath
        If Not FSO.FolderExists(strParentPath) Then
            If FSO.FileExists(strParentPath) Then
                ReportError "No path exists", _
                            "Path", strParentPath
                '                Err.Raise NoPathExists, _
                '                          RoutineName, _
                '                          NoPathExistsDescription & strParentPath
            Else
                FSO.CreateFolder (strParentPath)
            End If
        End If
    End If
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' EnsurePath

Private Sub UpdateConfigFile(ByVal config As clsConfiguration)

    ' This routine writes the configuration file
    ' Version 1.0.1
    ' Put this routine into normal form
    
    Const RoutineName As String = Module_Name & "UpdateConfigFile"
    On Error GoTo ErrorHandler
    
    Dim cPerf As PerformanceClass
    If gbDebug(RoutineName) Then
        Set cPerf = New PerformanceClass
        cPerf.SetRoutine RoutineName
    End If
    
    config.WriteToProjectConfigFile

    MsgBox "Configuration file was successfully updated. File:" & _
           vbCrLf & _
           config.ConfigFile & _
           vbCrLf & vbCrLf & _
           "Please review the file with a text editor. " & _
           "For details see: https://github.com/spences10/VBA-IDE-Code-Export#the-configuration-file", , _
           "Configuration File Updated"
    
    '@Ignore LineLabelNotUsed
Done:
    Exit Sub
ErrorHandler:
    RaiseError Err.Number, Err.Source, RoutineName, Err.Description
End Sub                                          ' UpdateConfigFile

Public Sub LetGitFormCanceled(ByVal Canx As Boolean)
    pFormCanceled = Canx
End Sub                                          ' LetGitFormCanceled

Public Sub LetGitFormDelete(ByVal DeleteCheck As Boolean)
    If DeleteCheck Then
        Select Case MsgBox( _
               "Are you sure you want to delete all the code in this project when you export?", _
               vbYesNo Or vbExclamation, _
               "Delete All Code?")
        Case vbYes
            pFormDeleted = True
        Case vbNo
            pFormDeleted = False
        End Select
    End If
End Sub                                          ' LetGitFormDelete


