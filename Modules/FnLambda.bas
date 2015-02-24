Attribute VB_Name = "FnLambda"
' Fn Lambdas
' ----------
'
' An external module that gives the functions here the passing invokation mechanism that FnBuffer has in place.
'
' Basically, for all other modules defining lambdas that want to use FnBuffer's function buffers.
' This is the place to put it

Public Const MODULE_PREFIX As String = "FnLambda."

' ## Functional Interfaces

'# Curries a function
Public Sub Curry_(Args As Variant)
    Dim MethodName As String, PreArgs As Variant, CurArgs As Variant, TotalArgs As Variant
    MethodName = Args(0)
    PreArgs = Args(1)
    CurArgs = Args(2)
    TotalArgs = FnArrayUtil.Chain(Array(PreArgs, CurArgs))
    Fn.Result = Fn.Invoke(MethodName, TotalArgs)
End Sub

'# Composes functions together
Public Sub Compose_(Args As Variant)
    Dim MethodNames As Variant, AccRes As Variant, MIndex As Long, InitArgs As Variant, MethodName As String
    MethodNames = Args(0)
    ' No Args(1) for Compose
    InitArgs = Args(2)
    
    AccRes = Fn.Invoke(ArrayUtil.Last(MethodNames), InitArgs)
    For MIndex = UBound(MethodNames) - 1 To LBound(MethodNames) Step -1
        MethodName = MethodNames(MIndex)
        AccRes = Fn.InvokeOneArg(MethodName, AccRes)
    Next
    Fn.Result = AccRes
End Sub

'# (Re)invokes a function with predefined arguments
Public Sub Reinvoke_(Args As Variant)
    Dim MethodName As String, PreArgs As Variant
    MethodName = Args(2)(0)
    PreArgs = Args(1)
    
    Fn.Result = Fn.Invoke(MethodName, PreArgs)
End Sub

'# Turns a function to an one argument function
Public Function Lambda_(Args As Variant)
    Dim MethodName As String, PreArgs As Variant, CurArgs As Variant
    MethodName = Args(0)
    CurArgs = Args(2)(0)
    Fn.Result = Fn.Invoke(MethodName, CurArgs)
End Function


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
    CurIndex = Fn.Closure
    Arr_ = Args(1)
    
    Fn.Result = Arr_(CurIndex)
    
    
    CurIndex = CurIndex + 1
    If CurIndex > UBound(Arr_) Then _
        CurIndex = LBound(Arr_)
    
    Fn.Closure = CurIndex
End Sub
