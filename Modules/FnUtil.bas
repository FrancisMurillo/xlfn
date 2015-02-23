Attribute VB_Name = "FnUtil"
' Fn Utilities
' ------------
'
' A set of functions utilizing Fn
'
' These functions end with the suffix Fn signifying the create Fn lambdas

'# Returns a constant value whenever called
Public Function ConstantFn(Constant As Variant) As String
    ConstantFn = Fn.Curry("FnFunction.Identity_", Array(Constant))
End Function
