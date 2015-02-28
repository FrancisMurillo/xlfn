Attribute VB_Name = "ChipInfo"
Public Sub WriteInfo()
    ChipReadInfo.References = Array( _
        "Microsoft Visual Basic for Applications Extensibility *", _
        "Microsoft Scripting Runtime")
    ChipReadInfo.Modules = Array( _
        "Fn", "FnBuffer", _
        "FnUtil", "FnArrayUtil", _
        "FnOperator", "FnFunction", "FnPredicate")
End Sub


