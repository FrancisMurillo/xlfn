Attribute VB_Name = "FnBlock"
' Functional Blocks
' =================

'# IF statement like LISP
'# Takes three no argument functions id
Public Sub If_(CondFs As String, TrueFs As String, FalseFs As String)
    Dim Cond As Boolean
    If Cond Then
        Fn.InvokeNoArgs TrueFs
    Else
        Fn.InvokeNoArgs FalseFs
    End If
End Sub

