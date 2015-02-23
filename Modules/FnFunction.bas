Attribute VB_Name = "FnFunction"
' Functional Functions
' --------------------
'
' A set of one argument functions for you

' ## Function Constants
'
' These constants are to aid typing with the intellisense

Public Const METHOD_PREFIX As String = "FnFunction."

Public Const Identity As String = METHOD_PREFIX & "Identity_"

' ## Generic Functions
'
' Functions doing somethings random or useful

'# Returns the argument
Public Function Identity_(Arg As Variant)
    Fn.Result = Arg
End Function



' ## Boolean Functions
'
' Boolean functions

'# Returns the inverse boolean value
Public Function Not_(Arg As Boolean)
    Fn.Result = Not Arg
End Function

' ## Mathematical Functions
'
' Mathematical functions

'# Returns the multiplicative inverse of the argument
Public Function Reciprocal_(Arg As Variant)
    Fn.Result = 1 / Arg
End Function

'# Returns the additive inverse of the argument
Public Function Negative_(Arg As Variant)
    Fn.Result = -Arg
End Function


