Attribute VB_Name = "ChipInit"
'' This module is to download or install all the Chip modules properly
'' This is self-contained of all the functions needed to run everything

'===========================
'Configurations
'===========================
'# This HTTP URL is where the Chip Workbook is stored
Public Const REPO_URL As String = "http://github.com/FrancisMurillo/xlchip/raw/master/xlchip-RELEASE.xlsm"
Public Const DEPENDENCY_LIST As String = "Microsoft Visual Basic for Applications Extensibility *;Microsoft Scripting Runtime"
Public Const LIST_DELIMITER As String = ";"


'===========================
'Main Functions
'===========================

'# Install Chip by downloading the stable release file from the repository
'# and copying the required modules
Public Sub InstallChipFromRepo()
On Error GoTo ErrHandler
    ClearScreen

    Debug.Print "Install Chip From Repository"
    Debug.Print "=============================="
    
    ' Check dependencies
    If Not PrecheckDependencies Then Exit Sub
    
    Dim Path As String
    Debug.Print "Downloading Chip from " & REPO_URL
    Path = DownloadFile ' Download file using the default settings
    
    Debug.Print "Installing Chip"
    InstallChip Path ' Install Chip
    
    Debug.Print "Installation success"
Cleanup:
    If Path <> "" Then
        Debug.Print "Removing temporary file " & Path
        DeleteFile Path
    End If
    Exit Sub
ErrHandler:
    Debug.Print _
        "Whoops! There was an error in loading the file. " & _
        "Make sure you selected the URL is correctly pointed to a Chip workbook and that you have an connection."
    Resume Cleanup
End Sub

'# Installs Chip via browse dialog, if the person has the book saved then it makes it easier.
Public Sub InstallChipLocally()
On Error GoTo ErrHandler
    ClearScreen
    Debug.Print "Install Chip Locally"
    Debug.Print "=============================="

    ' Check dependencies
    If Not PrecheckDependencies Then Exit Sub
    
    Dim Path As String
    Debug.Print "Select a Chip workbook"
    Path = BrowseFile
    If Path = "False" Then
        Debug.Print "No file was selected. Cancel installation"
        Exit Sub ' None was selected
    End If
    Debug.Print "Path: " & Path
    
    Debug.Print "Installing Chip"
    InstallChip Path ' Install Chip
    
    Debug.Print "Installation success"
    Exit Sub
ErrHandler:
    Debug.Print _
        "Whoops! There was an error in loading the file. " & _
        "Make sure you selected a Chip workbook."
End Sub

'# Removes Chip from the workbook
Public Sub UninstallChip()
On Error GoTo ErrHandler
    ClearScreen
    Debug.Print "Uninstalling Chip"
    Debug.Print "=============================="

    RemoveChip ActiveWorkbook
    
    Debug.Print "Uninstallation success"
    Exit Sub
ErrHandler:
    Debug.Print _
        "Whoops! There was an error removing Chip " & _
        "Try to remove the remaining Chip modules instead."
End Sub

'===========================
'Internal Functions
'===========================

'# Checks if the workbook can execute the installation process
Private Function PrecheckDependencies() As Boolean
    PrecheckDependencies = True

    If ActiveWorkbook.Path = "" Then
        Debug.Print "Save the workbook first before installing modules."
        PrecheckDependencies = False
        
        Exit Function
    End If

    Dim Dependencies As Variant
    Dependencies = Split(DEPENDENCY_LIST, LIST_DELIMITER)
    Debug.Print "Checking dependencies"
    If Not CheckDependencies(Dependencies) Then
        Debug.Print "One or more of the depedencies are not included. Make sure they are and installing again."
        Debug.Print "Required References:"
        For Each Depedency In Dependencies
            Debug.Print "# " & Depedency
        Next
        PrecheckDependencies = False
    End If
End Function

'# This removes all Chip* modules except this one
Private Sub RemoveChip(ChipBook As Workbook, _
        Optional Verbose As Boolean = True)
    Dim CurProj As VBProject, Modules As Variant, Module As Variant
    Set CurProj = ActiveWorkbook.VBProject
    Modules = ListWorkbookModuleObjects(ChipBook)
    If Verbose Then Debug.Print "Removing Chip Modules:"
    For Each Module In Modules
        If Module.Name Like "Chip*" And Module.Name <> "ChipInit" Then
            If Verbose Then Debug.Print "-- " & Module.Name
            CurProj.VBComponents.Remove Module
        End If
    Next
End Sub

'# This copies the modules from the Chip workbook
'# The last core function
'@ Exception: Propagate
Private Sub InstallChip(ChipBookPath As String, _
        Optional Verbose As Boolean = True)
On Error GoTo ErrHandler:
    ' Get all the modules from the Chip workbook
    Dim CurBook As Workbook, ChipBook As Workbook, CurProj As VBProject
    Dim Modules As Variant, Module As Variant, TempPath As String, NewModule As VBComponent
    If Verbose Then Debug.Print "Opening Chip book"
    Set CurBook = ActiveWorkbook
    Set CurProj = CurBook.VBProject
    Set ChipBook = Workbooks.Open(ChipBookPath, ReadOnly:=True)
    
    If Verbose Then Debug.Print "Installing modules:"
    Modules = ListWorkbookModuleObjects(ChipBook)
    For Each Module In Modules ' Get all Chip* modules
        If Module.Name Like "Chip*" Then
            If Module.Name <> "ChipInit" Then ' Ignore this module
                If HasModule(Module.Name, CurBook) Then
                    DeleteModule Module.Name, CurBook
                    If Verbose Then Debug.Print "+- " & Module.Name & "(Updated)"
                Else
                    If Verbose Then Debug.Print "++ " & Module.Name
                End If
            
                TempPath = "~" & Format(Now(), "yyyymmddhhmmss") & "mod"
                
                Module.Export TempPath
                Set NewModule = CurProj.VBComponents.Import(TempPath)
                
                DeleteFile TempPath
            End If
        End If
    Next
ErrHandler:
    If Err.Number <> 0 Then
        If Verbose Then Debug.Print _
            "Whoops! There was an error using the Chip book. " & _
            "Make sure the conditions are good for opening a workbook"
    End If
CloseBook:
    If Verbose Then Debug.Print "Closing Chip book"
    DoEvents
    ChipBook.Close SaveChanges:=False
    Exit Sub
End Sub

'# This checks if the VB Project has the required references to run the code
'@ Param: Dependencies > A zero string array of dependencies
'@ Exception: Propagate
Public Function CheckDependencies(Dependencies As Variant) As Boolean
    Dim References As Variant
    References = ListProjectReferences
        
    Dim Depedency As Variant, Reference As Variant, IsFound As Boolean
    For Each Dependency In Dependencies
        IsFound = False
        For Each Reference In References
            IsFound = Reference Like Dependency
            If IsFound Then Exit For
        Next
        If Not IsFound Then
            CheckDependencies = False
            Exit Function
        End If
    Next
    CheckDependencies = True
End Function


'===========================
'Helper Functions
'===========================

'# Clears the intermediate screen
Public Sub ClearScreen()
    Application.SendKeys "^g ^a {DEL}"
End Sub

'# Removes a module whether it exists or not
'# Used in making sure there are no duplicate modules
Public Sub DeleteModule(ModuleName As String, Book As Workbook)
On Error Resume Next
    Dim CurProj As VBProject, Module As VBComponent
    Set CurProj = Book.VBProject
    Set Module = CurProj.VBComponents(ModuleName)
    CurProj.VBComponents.Remove Module
    DoEvents
    Err.Clear
End Sub

'# Checks if an module exists
Public Function HasModule(ModuleName As String, Book As Workbook) As Boolean
On Error Resume Next
    HasModule = False
    HasModule = Not Book.VBProject.VBComponents(ModuleName) Is Nothing  ' This fails if the module does not exists thus defaulting to False
    Err.Clear
End Function

'# Lists the modules of an workbook
'# Primarily used to get all Chip modules
'@ Return: An array of VB Components
Public Function ListWorkbookModuleObjects(Book As Workbook) As Variant
    Dim Comp As VBComponent, Modules As Variant, Index As Long
    Modules = Array()
    ReDim Modules(0 To Book.VBProject.VBComponents.Count - 1)
    For Each Comp In Book.VBProject.VBComponents
        Set Modules(Index) = Comp
        Index = Index + 1
    Next
    ListWorkbookModuleObjects = Modules
End Function

'# This browses a file using the Open File Dialog
'# Primarily used to open a macro enabled file
'@ Return: The absolute path of the selected file, an "False" if none was selected
Public Function BrowseFile() As String
    BrowseFile = Application.GetOpenFilename _
    (Title:="Please choose a file to open", _
        FileFilter:="Excel Macro Enabled Files *.xlsm (*.xlsm),")
End Function

'# This downloads a file from the internet using the HTTP GET method
'# This is primarily used for downloading a binary file or the workbook repo needed
'! Taken from a site, modified to my use
'@ Return: The absolute path of the downloaded file, if path was not provided else the path itself
Public Function DownloadFile(Optional URL As String = REPO_URL, Optional Path As String = "")
    If Path = "" Then ' Create pseudo unique path
        Path = ActiveWorkbook.Path & Application.PathSeparator & "~" & Format(Now(), "yyyymmddhhmmss")
    End If

    Dim FileNum As Long
    Dim FileData() As Byte
    Dim MyFile As String
    Dim WHTTP As Object
    
    On Error Resume Next
        Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5")
        If Err.Number <> 0 Then
            Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5.1")
        End If
    On Error GoTo 0
    
    WHTTP.Open "GET", URL, False
    WHTTP.Send
    FileData = WHTTP.responseBody
    Set WHTTP = Nothing
    
    FileNum = FreeFile
    Open Path For Binary Access Write As #FileNum
        Put #FileNum, 1, FileData
    Close #FileNum
    
    DownloadFile = Path
    Exit Function
End Function

'# Deletes a file forcibly, it does not check whether it is a folder or the path does not exists
'# This is used to delete a temp file whether it still exists or not
Public Sub DeleteFile(FilePath As String)
    With New FileSystemObject
        If .FileExists(FilePath) Then
            .DeleteFile FilePath
        End If
    End With
End Sub

'# This returns an string array of the references used in this VBA Project
'# The strings are the name of the references, not the filename or path
'@ Return: A zero-based array of strings
Public Function ListProjectReferences() As Variant
    Dim ReferenceLength As Integer, Index As Long
    Dim References As Variant
    
    ReferenceLength = ActiveWorkbook.VBProject.References.Count
    If ReferenceLength = 0 Then
        ListProjectReferences = Array()
        Exit Function
    End If
    
    References = Array()
    ReDim References(1 To ReferenceLength)
    For Index = 1 To ActiveWorkbook.VBProject.References.Count
        With ActiveWorkbook.VBProject.References.Item(Index)
            References(Index) = .Description
        End With
    Next
    
    ReDim Preserve References(0 To ReferenceLength - 1)
    ListProjectReferences = References
End Function

'# This returns an array of project references, the objects themselves for use
'# This is used for setting up the test workbook to have the correct references
'@ Return: A zero-based array of references
Public Function ListProjectReferenceObjects() As Variant
    Dim ReferenceLength As Integer, Index As Long
    Dim References As Variant
    
    ReferenceLength = ActiveWorkbook.VBProject.References.Count
    If ReferenceLength = 0 Then
        ListProjectReferences = Array()
        Exit Function
    End If
    
    References = Array()
    ReDim References(1 To ReferenceLength)
    For Index = 1 To ActiveWorkbook.VBProject.References.Count
        Set References(Index) = ActiveWorkbook.VBProject.References.Item(Index)
    Next
    
    ReDim Preserve References(0 To ReferenceLength - 1)
    ListProjectReferenceObjects = References
End Function

'# Checks if the refrence exists for a workbook given its name
Public Function HasReference(ReferenceName As String, Book As Workbook) As Boolean
On Error Resume Next
    HasReference = False
    HasReference = Not Book.VBProject.References(ReferenceName) Is Nothing
    Err.Clear
End Function
