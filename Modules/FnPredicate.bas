Attribute VB_Name = "FnPredicate"
' Functional Predicates
' ---------------------
'
' Convenient functional predicates for your own use
'

' ## Mathematical

'# Checks if a number is even
Public Sub IsEven(Val As Variant)
    Fn.Result = ((Val Mod 2) = 0)
End Sub

'# Checks if a number is odd
Public Sub IsOdd(Val As Variant)
    Fn.Result = ((Val Mod 2) = 1)
End Sub
