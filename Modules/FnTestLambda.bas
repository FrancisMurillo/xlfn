Attribute VB_Name = "FnTestLambda"
' Unit Test Functions
' -------------------
'
' These function are used in the unit testing, not used in production
' As stated, these functions set the result variable instead

'# Turns a number to its negative
Public Sub Negative_(Val As Long)
    Fn.Result = -1 * Val
End Sub

'# Adds a prefix to a string, 'Pre: ' prefix
Public Sub Prefix_(Val As String)
    Fn.Result = "Pre: " & Val
End Sub

'# Just wraps the value into an array
Public Sub WrapArray_(Val As Variant)
    Fn.Result = Array(Val)
End Sub

'# Accepts only 2
Public Sub IsTwo_(Val As Long)
    Fn.Result = (Val = 2)
End Sub

'# Accepts only Francis
Public Sub IsFrancis_(Val As String)
    Fn.Result = (Val = "Francis")
End Sub

'# Accepts all, a default filter for All
Public Sub True_(Val As Variant)
    Fn.Result = True
End Sub

'# Adds two numbers
Public Sub Add_(Acc As Long, Elem As Long)
    Fn.Result = Acc + Elem
End Sub

'# Concats strings
Public Sub Concat_(Acc As String, Elem As String)
    Fn.Result = Acc & Elem
End Sub

'# Makes Empty elements countable
Public Sub EmptyCount_(Acc As Long, Elem As Variant)
    Fn.Result = Acc + IIf(IsEmpty(Elem), 1, 0)
End Sub

'# Random tripet formula
Public Sub Formula_(Tuple As Variant)
    Dim Product As Long, Elem_ As Variant
    Product = 1
    For Each Elem_ In Tuple
        Product = Product * Elem_
    Next
    Fn.Result = Product
End Sub

'# A quick triple sum
Public Sub TripletSum_(A As Variant, B As Variant, C As Variant)
    Fn.Result = A + B + C
End Sub

'# Remove all letter a's
Public Sub RemoveA_(Val As String)
    Fn.Result = Replace(Val, "a", "")
End Sub

'# Remove all letter i's
Public Sub RemoveI_(Val As String)
    Fn.Result = Replace(Val, "i", "")
End Sub

'# Go to uppercase
Public Sub ToUppercase_(Val As String)
    Fn.Result = UCase(Val)
End Sub

'# Operator and arguments
Public Sub OperatorApply_(LVal As Variant, RVal As Variant, OperatorName As String)
    Fn.Result = Fn.InvokeTwoArg(OperatorName, LVal, RVal)
End Sub


'# Add one item to an collection
Public Sub AddOneToCollection_(Col As Collection)
    Col.Add 1
    Set Fn.Result = Col
End Sub

'# Double the collection with the same elements
Public Sub DoubleCollection_(Col As Collection)
    Dim Col_ As New Collection, Elem As Variant
    For Each Elem In Col
        Col_.Add Elem
    Next
    ' Twice
    For Each Elem In Col
        Col_.Add Elem
    Next
    
    Set Fn.Result = Col_
End Sub

Public Sub JoinCollection_(LCol As Collection, RCol As Collection)
    Dim Col_ As New Collection, Elem_ As Variant
    For Each Elem_ In LCol
        Col_.Add Elem_
    Next
    For Each Elem_ In RCol
        Col_.Add Elem_
    Next
    
    Set Fn.Result = Col_
End Sub
