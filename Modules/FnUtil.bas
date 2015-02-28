Attribute VB_Name = "FnUtil"
' Fn Utilities
' ------------
'
' A set of functions utilizing Fn
'
' These functions end with the suffix Fn signifying the create Fn lambdas

' ## Filtering Utilities

'# Decorate the function with a not operator
Public Function WrapNot_(BoolFn As String, _
                    Optional ClosureArgs As Variant = Empty) As String
    WrapNot_ = Fn.Decorate(FnFunction.Negative_Fn, BoolFn, ClosureArgs:=ClosureArgs)
End Function
