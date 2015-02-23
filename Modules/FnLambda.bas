Attribute VB_Name = "FnLambda"
' Fn Lambdas
' ----------
'
' An external module that gives the functions here the passing invokation mechanism that FnBuffer has in place.
'
' Basically, for all other modules defining lambdas that want to use FnBuffer's function buffers.
' This is the place to put it

Public Const MODULE_PREFIX As String = "FnLambda."

' ## Iterator Lambdas

'# Constantly returns one value
Private Sub Constant_Lambda(Args As Variant)
    Dim Val As Variant
    Val = Args(1)
    Fn.Result = Val
End Sub

'# Cycles through an array
Private Sub Cycle_Lambda(Args As Variant)
    Dim BufferIndex As Long, CurIndex As Long, Arr_ As Variant
    BufferIndex = Args(4)
    CurIndex = Args(3)
    Arr_ = Args(1)
    
    Fn.Result = Arr_(CurIndex)
    
    CurIndex = CurIndex + 1
    If CurIndex > UBound(Arr_) Then _
        CurIndex = LBound(Arr_)
    
    FnBuffer.SetClosureBufferArgs CurIndex, BufferIndex
End Sub
