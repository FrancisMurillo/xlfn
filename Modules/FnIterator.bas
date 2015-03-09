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
Public Function Next_(IteratorFp As Variant) As Variant
    Assign_ Next_, Fn.InvokeNoArgs(IteratorFp)
End Function

'# Returns an array representing the number of times an iterator is repeated
'R Zero Base
Public Function Iterate(IteratorFp As Variant, Count As Long) As Variant
    If Count <= 0 Then
        Iterate = Array()
    Else
        Dim Iterate_ As Variant, Index As Long
        Iterate_ = ArrayUtil.CreateWithSize(Count)
        For Index = LBound(Iterate_) To UBound(Iterate_)
            Assign_ Iterate_(Index), Next_(IteratorFp)
        Next
        Iterate = Iterate_
    End If
End Function

Public Function Random(Optional Start_ As Long = 0, Optional End_ As Long = 1000, Optional Seed As Long = 0) As Variant
    If Start_ >= End_ Then _
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "Random Start cannot be less than End range"

    Randomize IIf(Seed = 0, Now, Seed)
    Random = Fn.CreateLambda(RANDOM_METHOD, Empty, Array(Start_, End_), Empty)
End Function
Private Sub Random_Fn(Optional Args As Variant = Empty)
    Dim Val As Long, Start_ As Long, End_ As Long
    Start_ = Fn.PreArgs(0)
    End_ = Fn.PreArgs(1)
    Val = Abs(End_ - Start_) * Rnd() + Start_
    Fn.AssignResult_ Val
End Sub

'# Returns a constant value
Public Function Constant(Val As Variant) As Variant
    Constant = Fn.CreateLambda(CONSTANT_METHOD, Empty, Val, Empty)
End Function
Private Sub Constant_Fn(Optional Args As Variant = Empty)
    Dim Val As Variant
    Assign_ Val, Fn.PreArgs
    Fn.AssignResult_ Val
End Sub


'# Returns a function string that loops through an array ad infinitum
'# If Arr is empty, this defaults to constat Empty
Public Function Cycle(Arr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Arr) Then
        Cycle = Constant(Empty)
    Else
        Cycle = Fn.CreateLambda(CYCLE_METHOD, Empty, Arr, LBound(Arr))
    End If
End Function
Private Sub Cycle_Fn(Optional Args As Variant = Empty)
    Dim BufferIndex As Long, CurIndex As Long, Arr_ As Variant
    CurIndex = Fn.Closure
    Assign_ Arr_, Fn.PreArgs
    
    Fn.AssignResult_ Arr_(CurIndex)
        
    CurIndex = CurIndex + 1
    If CurIndex > UBound(Arr_) Then _
        CurIndex = LBound(Arr_)
    
    Fn.AssignClosure_ CurIndex
    ' Assign_ Fn.Closure, CurIndex
End Sub


' ## Utility function
Private Sub Assign_(ByRef Assignee As Variant, ByVal Assigned As Variant)
    If IsObject(Assigned) Then
        Set Assignee = Assigned
    Else
        Assignee = Assigned
    End If
End Sub
