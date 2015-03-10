Attribute VB_Name = "FnUtil"
' Fn Utilities
' ------------
'
' A set of functions utilizing Fn
'
' These functions end with the suffix Fn signifying the create Fn lambdas

' ## Constants

Private Const MEMOIZE_METHOD As String = "FnUtil.Memoize_Fn"
Private Const HASH_DEFAULT_METHOD As String = "FnUtil.HashDefault_Fn"
Private Const TIMEIT_METHOD As String = "FnUtil.Timeit_Fn"

' ## Filtering Utilities

'# Decorate the function with a not operator
Public Function WrapNot_(BoolFn As String, _
                    Optional ClosureArgs As Variant = Empty) As String
    WrapNot_ = Fn.Decorate(FnFunction.Negative_Fn, BoolFn, ClosureArgs:=ClosureArgs)
End Function

'# Memoizes a function with a default hashing function
Public Function Memoize(Fp, Optional HashFp As Variant = Empty) As Variant
    If IsEmpty(HashFp) Then
        HashFp = HASH_DEFAULT_METHOD
    End If

    Memoize = Fn.CreateLambda(MEMOIZE_METHOD, Fp, Array(New Dictionary, HashFp), Empty)
End Function
Private Sub HashDefault_Fn(Args As Variant)
    If IsEmpty(Args) Then
        Fn.Result = CStr(Empty)
    Else
        Fn.Result = CStr(Join(Args, "*|*"))
    End If
End Sub
Private Sub Memoize_Fn(Optional Args As Variant = Empty)
    If IsMissing(Args) Then _
        Args = Empty

    Dim MemoCol As Dictionary, HashKey As Variant, HashFp As Variant
    Set MemoCol = Fn.PreArgs(0)
    HashFp = Fn.PreArgs(1)
    HashKey = Fn.InvokeOneArg(HashFp, Args)
    
    If MemoCol.Exists(HashKey) Then
        Fn.Result = MemoCol.Item(HashKey)
    Else
        Dim Res As Variant
        Assign_ Res, Fn.Invoke(Fn.NextFp, Args)
        Fn.AssignResult_ Res
        MemoCol.Add HashKey, Res
    End If
End Sub

'# Inspired from Python's timeit, measures a function execution
Public Function Timeit(Fp As Variant) As Variant
    Timeit = Fn.CreateLambda(TIMEIT_METHOD, Fp, Empty, Empty)
End Function
Private Function Timeit_Fn(Args As Variant) As Variant
    Dim Start_ As Long, Res_ As Variant
    Start_ = Timer
    Assign_ Res_, Fn.Invoke(Fn.NextFp, Args)
    Fn.Result = Array(Res_, Timer - Start_)
End Function

' ## Utility function
Private Sub Assign_(ByRef Assignee As Variant, ByVal Assigned As Variant)
    If IsObject(Assigned) Then
        Set Assignee = Assigned
    Else
        Assignee = Assigned
    End If
End Sub
