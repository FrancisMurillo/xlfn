Attribute VB_Name = "FnPredicate"
' Functional Predicates
' ---------------------
'
' Convenient functional predicates for your own use
'

' ## Generic

'# Checks if a value is empty, basically wraps IsEmpty as a lambda
Public Sub IsEmpty_(Val As Variant)
    IsEmpty_ = IsEmpty(Val)
End Sub



' ## Mathematical

'# Checks if a number is even
Public Sub IsEven_(Val As Variant)
    Fn.Result = ((Val Mod 2) = 0)
End Sub

'# Checks if a number is odd
Public Sub IsOdd_(Val As Variant)
    Fn.Result = ((Val Mod 2) = 1)
End Sub
