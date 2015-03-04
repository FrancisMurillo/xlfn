Attribute VB_Name = "FnIterator"
' Fn Iterator Utility
' -------------------
'
' This is a toy module to create iterators in this programming language
'

' ## Constants
'
' Error constants
Private Const ERR_SOURCE As String = "FnIterator"
Private Const ERR_OFFSET As Long = 2200

Private Const Constant_Lambda As String = "Constant_Lambda"
Private Const Cycle_Lambda As String = "Cycle_Lambda"
Private Const Random_Lambda As String = "Random_Lambda"

' Lambda Function Name Constants
Private Const MODULE_PREFIX As String = "FnIterator."

Private Const CONSTANT_METHOD As String = MODULE_PREFIX & "Constant" & Fn.LAMBDA_SUFFIX
Private Const RANDOM_METHOD As String = MODULE_PREFIX & "Random" & Fn.LAMBDA_SUFFIX
Private Const CYCLE_METHOD As String = MODULE_PREFIX & "Cycle" & Fn.LAMBDA_SUFFIX


' ## Iterator Functions

'# An alias to geet the next value for an iterator
Public Function Next_(IteratorFn As String) As Variant
    Next_ = Fn.InvokeNoArgs(IteratorFn)
End Function

'# Returns an array representing the number of times an iterator is repeated
'R Zero Base
Public Function Iterate(IteratorFn As String, Count As Long) As Variant
    If Count <= 0 Then
        Iterate = Array()
    Else
        Dim Iterate_ As Variant, Index As Long
        Iterate_ = ArrayUtil.CreateWithSize(Count)
        For Index = LBound(Iterate_) To UBound(Iterate_)
            Iterate_(Index) = Next_(IteratorFn)
        Next
        Iterate = Iterate_
    End If
End Function


Public Function Random(Optional Start_ As Long = 0, Optional End_ As Long = 1000, Optional Seed As Long = 0)
    If Start_ >= End_ Then _
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "Random Start cannot be less than End range"

    Randomize IIf(Seed = 0, Now, Seed)
    Random = FnBuffer.GenerateBufferLambda(RANDOM_METHOD, Empty, Array(Start_, End_), Empty)
End Function
Private Sub Random_Fn(Args As Variant)
    Dim Val As Long, Start_ As Long, End_ As Long
    Start_ = Args(1)(0)
    End_ = Args(1)(1)
    Val = Abs(End_ - Start_) * Rnd() + Start_
    Fn.Result = Val
End Sub

'# Returns a constant value
Public Function Constant(Val As Variant) As String
    Constant = FnBuffer.GenerateBufferLambda(CONSTANT_METHOD, Empty, Val, Empty)
End Function
Private Sub Constant_Fn(Args As Variant)
    Dim Val As Variant
    Val = Args(1)
    Fn.Result = Val
End Sub


'# Returns a function string that loops through an array ad infinitum
'# If Arr is empty, this defaults to constat Empty
Public Function Cycle(Arr As Variant) As String
    If ArrayUtil.IsEmptyArray(Arr) Then
        Cycle = Constant(Empty)
    Else
        Cycle = FnBuffer.GenerateBufferLambda(CYCLE_METHOD, Empty, Arr, LBound(Arr))
    End If
End Function
Private Sub Cycle_Fn(Args As Variant)
    Dim BufferIndex As Long, CurIndex As Long, Arr_ As Variant
    CurIndex = Fn.Closure
    Arr_ = Args(1)
    
    Fn.Result = Arr_(CurIndex)
        
    CurIndex = CurIndex + 1
    If CurIndex > UBound(Arr_) Then _
        CurIndex = LBound(Arr_)
    
    Fn.Closure = CurIndex
End Sub



