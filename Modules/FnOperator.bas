Attribute VB_Name = "FnOperator"
' Functional Operators
' --------------------
'
' These set of operator function are implementation of the common VB operators
' Just a convenience for Reduce_ or others
'
' There is just one caveat: the type conversion of variant to a concrete type
' Just be wary of converting the values since you might not get the expected type
' But since we're using variant, this might hold little value

' ## Mathematical Operators
'
' Math operators
' The only caveat is that there is no type signature, so the returning result is dependent on the operators evaluation
' So adding Long and Decimal might not produce the correct type you want

'# Adds two numberss
Public Sub Add_(LVal As Variant, RVal As Variant)
    Fn.Result = (LVal + RVal)
End Sub

'# Multiples two numbers
Public Sub Multiply_(LVal As Variant, RVal As Variant)
    Fn.Result = (LVal * RVal)
End Sub

' ## Logical Operators
'
' And or
' Every value is typed as boolean so the return type is boolean

'# ORs two values
Public Sub Or_(LVal As Boolean, RVal As Boolean)
    Fn.Result = (LVal Or RVal)
End Sub

'# ANDs two values
Public Sub And_(LVal As Boolean, RVal As Boolean)
    Fn.Result = (LVal And RVal)
End Sub


' ## String Operator
'
' String operators

'# Concatenates two strings
Public Sub Concat_(LStr As String, RStr As String)
    Fn.Result = (LStr & RStr)
End Sub

'# Checks if one string is like the other
Public Sub Like_(LStr As String, RPat As String)
    Fn.Result = LStr Like RPat
End Sub

