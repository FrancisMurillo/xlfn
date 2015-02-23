Attribute VB_Name = "WheatConfig"
'# Wheat Configuration File
'# My answer to persistence configuration of an Excel file
'# As a side note this file should not be exported or imported since this is a local configuration

'# Currently right now, I only need a few options. So I'll stick with it
'# Might expand as an option module

'# PROJECT REPO
'# The name of the project folder, an absolute or relative path.
Public Const PROJECT_REPO As String = "xlfn-src"

'# EXPORT Options
Public Const SHOW_EXPORTED_MODULES As Boolean = True
Public Const SHOW_IGNORED_MODULES As Boolean = True
Public Const SHOW_IGNORED_EXCEPT_MODULES As Boolean = True

Public IgnoreExportModules As Variant ' Modules you want to ignore when exporting
Public IgnoreExceptExportModules As Variant ' Modules you want to export regardless when ignored

'# IMPORT Options
Public Const SHOW_IMPORTED_MODULES As Boolean = True
Public Const SHOW_PASSED_MODULES As Boolean = True
Public Const SHOW_PASSED_EXCEPT_MODULES As Boolean = True

Public PassImportModules As Variant ' Modules you want to ignore during import
Public PassExceptImportModules As Variant ' Modules that are exempt from the ignore import filter

Public Sub InitializeVariables()
    ' Sample modules to ignore, a reasonable default is provided
    IgnoreExportModules = Array( _
        "Chip*", "Vase*", "Wheat*", _
        "Sheet*", "ThisWorkbook", _
        "*_", _
        "Sandbox", "Control")
    IgnoreExceptExportModules = Array( _
        "ChipInfo", "ChipInit", "WheatConfig")
    
    ' Same restriction as exporting
    ' Modify this when to your specific needs
    PassImportModules = IgnoreExportModules
    PassExceptImportModules = IgnoreExceptExportModules
End Sub


