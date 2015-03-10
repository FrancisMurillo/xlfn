Attribute VB_Name = "Fn"
' Functional Programming: Fn
' ------------------
'
' This module provides a mechanism to provide pseudo lambda in VBA.
'
' # Module Definition
'
' Since we don't have first class functions, their names could be their pointers instead.
' By using Application.Run, we can invoke the function.
' Sadly, Application.Run has some limitation like not having a return value;
' This is remedied by setting the property Result(should have been Return but its already a keyword)
' with the value of the function.
'
' So if you have a function in a module, MyModule, like so...
'
' Public Function MyFunc(MyArg as Variant) As String
'   MyFunc = Str(MyArg)
' End Function
'
' Under the definition of pseudo lambda it will be...
'
' Public Sub MyFunc(MyArg as Variant)
'   Assign_ Fn.Result, Str(MyArg)
' End Sub
'
' This pseudo function can be invoked by...
'
' Fn.Invoke("MyModule.MyFunc", Array(MyArg))
'
' Not a whole lot of difference except for Function being Sub
' and the return mechanism, which I say might be better than writing the function name all the time,
' and the invokation mechanism, something of a necessary evil
'
' So what does wrapping or fitting the function to lambdas get you?
' It gives you the ability to invoke them with cool functional methods like Map, Reduce and Filter.
' If you tasted the first kiss of functional programming, this is a little drop of heaven.
' This is better than doing procedural and crappy OO.
'
' Modules with the Fn prefix(except this one) utilize this mechanism. If you are making your own, you should too for convention.
' Methods utilizing the Fn.Invoke should be placed in the modules with Fn and end with undersocre for convention.
' Not a requirement, but helps with familiarity although I use this convention as well to avoid naming conflicts.
'
' Word of warning, this mechanism trades flexibility for performance.
' So when using this for performance critical aspects, take of your gloves and get your hands dirty.
'
' So be mindful when and where to use this. Such is the way of the programmer.

' ## Consants
'
' Error constants
Private Const ERR_SOURCE As String = "Fn"
Private Const ERR_OFFSET As Long = 2000

' Lambda Constant
Public Const LAMBDA_SUFFIX As String = "_Fn"

' Lambda Function Name Constants
Private Const MODULE_PREFIX As String = "Fn."

Private Const CURRY_METHOD As String = MODULE_PREFIX & "Curry" & LAMBDA_SUFFIX
Private Const COMPOSE_METHOD As String = MODULE_PREFIX & "Compose" & LAMBDA_SUFFIX
Private Const WITH_ARGS_METHOD As String = MODULE_PREFIX & "WithArgs" & LAMBDA_SUFFIX
Private Const UNPACK_METHOD As String = MODULE_PREFIX & "Unpack" & LAMBDA_SUFFIX
Private Const DECORATE_METHOD As String = MODULE_PREFIX & "Decorate" & LAMBDA_SUFFIX

' ## Property

Private gResult As Variant
Private gNextFp As Variant
Private gPreArgs As Variant
Private gClosure As Variant
Private gThisFp As Variant


'# The Result property, place your result here. Write-only, that's what it's supposed to be.
Public Property Get Result()
    Assign_ Result, gResult
End Property
Public Property Let Result(Val As Variant)
    gResult = Val
End Property
Public Property Set Result(Val As Variant)
    Set gResult = Val
End Property

'# The Closure property for easier read and write
Public Property Let Closure(Val As Variant)
    gClosure = Val
End Property
Public Property Get Closure() As Variant
    Closure = gClosure
End Property

'# Buffer Index property


'# The arguments applied before the buffer function
Public Property Get PreArgs() As Variant
    Assign_ PreArgs, gPreArgs
End Property

'# The next method to invoke if there is one
Public Property Get NextFp() As Variant
    Assign_ NextFp, gNextFp
End Property

'# This returns the current calling method, thus allowing recursion
Public Property Get ThisFp() As Variant
    Assign_ ThisFp, gThisFp
End Property

' ## Functions

'# Invokes a function given its name and an array of arguments
'# This is achieved by using Application.Run and the concept of functions just have one argument
'# There is one subtle limitation, the maximum number of arguments. Due to Application.Run, the maximum number is 30.
'# Anything higher would result in an error
'P MethodName: The method to be invoked given its name. It should be the full name to be exact like "Fn.Invoke", not just "Invoke".
'P             You can just use the method name but you might run a function of the same name, so long name for safety.
'P Args: This is the arguments for the method wrapped in an array.
'P       This is also assumed to have base zero, but not a strict condition.
'P       The arguments are applied by order not by index, but make my our easier by using Array() to wrap the arguments
Public Function Invoke(ByRef MethodFp As Variant, Args As Variant) As Variant
On Error GoTo ErrHandler:
    Dim Args_ As Variant
    Args_ = ArrayUtil.AsNormalArray(Args)
    
    ' Reset variables
    gResult = Empty
    gClosure = Empty
    gBufferIndex = Empty
    gPreArgs = Empty
    gNextFn = Empty
    
    Assign_ gThisFp, MethodFp
            
    If IsCompositeFunction(MethodFp) Then
        Dim LambdaFs As String, Closure As Variant
        LambdaFs = MethodFp(0)
        Assign_ gNextFp, MethodFp(1)
        Assign_ gPreArgs, MethodFp(2)
        Assign_ gClosure, MethodFp(3)
        
        NonLeafInvokation LambdaFs, Args_
        
        Assign_ MethodFp(3), gClosure
    Else
        LeafInvokation CStr(MethodFp), Args_
    End If
    
    Assign_ Invoke, gResult
ErrHandler:
    If Err.Number = 1004 Then
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "The method " & MethodName & " does not exist"
    ElseIf Err.Number <> 0 Then
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, MethodName & " caused an error: " & Err.Description
    End If
End Function

'# Checks if it is a lambda function name
Public Function IsLambdaFunctionName(MethodName As String) As Boolean
    IsLambdaFunctionName = (Right(MethodName, Len(LAMBDA_SUFFIX)) = LAMBDA_SUFFIX)
End Function

Public Function IsBufferFunctionName(MethodName As String) As Boolean
    IsBufferFunctionName = (Left(MethodName, Len(FnBuffer.BUFFER_MAIN_METHOD)) = FnBuffer.BUFFER_MAIN_METHOD)
End Function

'# Checks if the array
Public Function IsCompositeFunction(Fc As Variant)
    IsCompositeFunction = False
    If IsArray(Fc) Then
        If UBound(Fc) = 3 Then
            IsCompositeFunction = True
        End If
    End If
End Function

'# Creates the pseudo lambda definition
Public Function CreateLambda(LambdaFs As String, InnerFp As Variant, PreArgs As Variant, ClosureVars As Variant) As Variant
    CreateLambda = Array(LambdaFs, InnerFp, PreArgs, ClosureVars)
End Function


'# Just passes the invokation
Public Sub NonLeafInvokation(MethodName As String, Args As Variant)
    If ArrayUtil.IsEmptyArray(Args) Then
        Application.Run MethodName
    Else
        Application.Run MethodName, Args
    End If
End Sub

'# The very application of an invokation
Public Sub LeafInvokation(MethodName As String, Args As Variant)
    Dim Size_ As Long
    Size_ = ArrayUtil.Size(Args)
    ' The long case of Application.Run, Python FTW
    Select Case Size_
        Case 0
            Application.Run MethodName
        Case 1
            Application.Run MethodName, Args(0)
        Case 2
            Application.Run MethodName, Args(0), Args(1)
        Case 3
            Application.Run MethodName, Args(0), Args(1), Args(2)
        Case 4
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3)
        Case 5
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4)
        Case 6
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5)
        Case 7
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6)
        Case 8
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7)
        Case 9
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8)
        Case 10
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9)
        Case 11
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10)
        Case 12
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11)
        Case 13
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12)
        Case 14
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13)
        Case 15
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14)
        Case 16
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15)
        Case 17
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16)
        Case 18
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17)
        Case 19
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18)
        Case 20
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19)
        Case 21
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20)
        Case 22
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21)
        Case 23
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22)
        Case 24
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23)
        Case 25
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24)
        Case 26
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25)
        Case 27
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26)
        Case 28
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27)
        Case 29
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27), Args(28)
        Case 30
            Application.Run MethodName, Args(0), Args(1), Args(2), Args(3), Args(4), Args(5), Args(6), Args(7), Args(8), Args(9), Args(10), Args(11), Args(12), Args(13), Args(14), Args(15), Args(16), Args(17), Args(18), Args(19), Args(20), Args(21), Args(22), Args(23), Args(24), Args(25), Args(26), Args(27), Args(28), Args(29)
        Case Else
            Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "Invoking " & MethodName & " with " & Size_ & " arguments exceeded the maximum number(30)"
    End Select
End Sub

'# Invokes a method without arguments
Public Function InvokeNoArgs(MethodFp As Variant)
    Assign_ InvokeNoArgs, Invoke(MethodFp, Array())
End Function

'# Invokes a method with one argument
Public Function InvokeOneArg(MethodFp As Variant, Arg As Variant)
    Assign_ InvokeOneArg, Invoke(MethodFp, Array(Arg))
End Function

'# Invokes a method with two arguments
Public Function InvokeTwoArg(MethodFp As Variant, Arg1 As Variant, Arg2 As Variant)
    Assign_ InvokeTwoArg, Invoke(MethodFp, Array(Arg1, Arg2))
End Function

'# Just a function to easily test the installation of Fn
Public Sub Hello()
    Debug.Print Fn.InvokeOneArg(FnFunction.Identity_Fn, "Hello Fn: The Pseudo Functional Programming Library for VBA")
End Sub

' ## Combinator Functions
'
' These functions combines functions basically.

'# Curries functions, returns the buffer name to be used by invoke
Public Function Curry(MethodFp As Variant, PreArgs As Variant) As Variant
    Curry = CreateLambda(CURRY_METHOD, MethodFp, PreArgs, Empty)
End Function
Private Sub Curry_Fn(Optional Args As Variant = Empty)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    AssignResult_ Fn.Invoke(Fn.NextFp, ArrayUtil.JoinArrays(Fn.PreArgs, Args))
End Sub

'# Combines several functions together, think of function composition here
Public Function Compose(MethodFps As Variant) As Variant
    Compose = CreateLambda(COMPOSE_METHOD, Empty, MethodFps, Empty)
    'Compose = FnBuffer.GenerateBufferLambda(COMPOSE_METHOD, Empty, MethodNames, Empty)
End Function
Private Sub Compose_Fn(Optional Args As Variant = Empty)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
            
    Dim MethodFps As Variant, AccRes As Variant, MIndex As Long, InitArgs As Variant, MethodFp As Variant
    MethodFps = Fn.PreArgs
    Assign_ AccRes, Fn.Invoke(ArrayUtil.Last(MethodFps), Args)
    For MIndex = UBound(MethodFps) - 1 To LBound(MethodFps) Step -1
        MethodFp = MethodFps(MIndex)
        Assign_ AccRes, Fn.InvokeOneArg(MethodFp, AccRes)
    Next
    AssignResult_ AccRes
End Sub

'# This is similar to curry but this functions more as a closure or a deferred executor
'# This function accepts a method name given predefined arguments
'# Primarily used to Map array of functions given arguments
'# This gives you the ability to put the function name as the parameter
Public Function WithArgs(Args As Variant) As Variant
    WithArgs = CreateLambda(WITH_ARGS_METHOD, Empty, Args, Empty)
End Function
'# (Re)invokes a function with predefined arguments
Private Sub WithArgs_Fn(Args As Variant)
    Dim MethodFp As Variant
    Assign_ MethodFp, Args(0)
    AssignResult_ Fn.Invoke(MethodFp, Fn.PreArgs)
End Sub


'# Wraps a function to accept an argument array instead of a plain argument
'# This is used basically wrapped multiple arguments to one, quite hard to explain
Public Function Unpack(MethodFp As Variant) As Variant
    Unpack = CreateLambda(UNPACK_METHOD, MethodFp, Empty, Empty)
End Function
Private Function Unpack_Fn(Args As Variant)
    AssignResult_ Fn.Invoke(Fn.NextFp, Args(0))
End Function

'# A shorter form of compose, decorate just handles one function
Public Function Decorate(WrappingFp As Variant, WrappedFp As Variant)
    Decorate = CreateLambda(DECORATE_METHOD, Empty, Array(WrappingFp, WrappedFp), Empty)
End Function
Private Function Decorate_Fn(Args As Variant)
    Dim WrappedFp As String, WrappingFp As String
    Assign_ WrappingFp, Fn.PreArgs(0)
    Assign_ WrappedFp, Fn.PreArgs(1)
    AssignResult_ Fn.InvokeOneArg(WrappingFp, Fn.Invoke(WrappedFp, Args))
End Function



' ## Utility function
Private Sub Assign_(ByRef Assignee As Variant, ByVal Assigned As Variant)
    If IsObject(Assigned) Then
        Set Assignee = Assigned
    Else
        Assignee = Assigned
    End If
End Sub

'# A set of generic assignment that supports objects and values without SET keyword
Public Sub AssignResult_(ByVal Assigned As Variant)
    Assign_ gResult, Assigned
End Sub
Public Sub AssignClosure_(ByVal Assigned As Variant)
    Assign_ gClosure, Assigned
End Sub
