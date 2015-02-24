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
'   Fn.Result = Str(MyArg)
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

Private Const BUFFER_MODULE As String = "FnBuffer"
Private Const LAMBDA_MODULE As String = "FnLambda"
Private Const BUFFER_PATTERN As String = BUFFER_MODULE & ".*"
Private Const LAMBDA_PATTERN As String = LAMBDA_MODULE & ".*"

' ## Property

Private gResult As Variant
Private gClosure As Variant
Private gBufferIndex As Long
Private gPreArgs As Variant

'# The Result property, place your result here. Write-only, that's what it's supposed to be.

Public Property Let Result(Val As Variant)
    gResult = Val
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
    PreArgs = gPreArgs
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
Public Function Invoke(MethodName As String, Args As Variant) As Variant
On Error GoTo ErrHandler:
    Dim Args_ As Variant
    Args_ = ArrayUtil.AsNormalArray(Args)
    
    ' Reset variables
    gResult = Empty
    gClosure = Empty
    gBufferIndex = Empty
    gPreArgs = Empty
    
    If MethodName Like BUFFER_PATTERN Then
        NonLeafInvokation MethodName, Args_
    ElseIf MethodName Like LAMBDA_PATTERN Then
        gClosure = Args_(3)
        gBufferIndex = Args_(4)
    
        NonLeafInvokation MethodName, Args_
    
        Call FnBuffer.SetClosureBufferArgs(gClosure, gBufferIndex)
    Else
        LeafInvokation MethodName, Args_
    End If
    
    Invoke = gResult
ErrHandler:
    If Err.Number = 1004 Then
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "The method " & MethodName & " does not exist"
    ElseIf Err.Number <> 0 Then
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, MethodName & " caused an error: " & Err.Description
    End If
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
Public Function InvokeNoArgs(MethodName As String)
    InvokeNoArgs = Invoke(MethodName, Array())
End Function

'# Invokes a method with one argument
Public Function InvokeOneArg(MethodName As String, Arg As Variant)
    InvokeOneArg = Invoke(MethodName, Array(Arg))
End Function

'# Invokes a method with two arguments
Public Function InvokeTwoArg(MethodName As String, Arg1 As Variant, Arg2 As Variant)
    InvokeTwoArg = Invoke(MethodName, Array(Arg1, Arg2))
End Function

'# Just a function to easily test the installation of Fn
Public Sub Hello()
    Debug.Print Fn.InvokeOneArg(FnFunction.Identity_Fn, "Hello Fn: The Pseudo Functional Programming Library for VBA")
End Sub

' ## Combinator Functions
'
' These functions combines functions basically.

'# Curries functions, returns the buffer name to be used by invoke
Public Function Curry(MethodName As String, PreArgs As Variant, _
                    Optional ClosureArgs As Variant = Empty) As String
    Curry = GenerateLambdaBufferDefinition(FnBuffer.CURRY_METHOD, MethodName, PreArgs, ClosureArgs)
End Function

'# Combines several functions together, think of function composition here
Public Function Compose(MethodNames As Variant, _
                    Optional ClosureArgs As Variant = Empty) As String
    Compose = GenerateLambdaBufferDefinition(FnBuffer.COMPOSE_METHOD, MethodNames, Empty, ClosureArgs)
End Function

'# This is similar to curry but this functions more as a closure or a deferred executor
'# This function accepts a method name given predefined arguments
'# Primarily used to Map array of functions given arguments
'# This gives you the ability to put the function name as the parameter
Public Function Reinvoke(Args As Variant, _
                    Optional ClosureArgs As Variant = Empty)
    Reinvoke = GenerateLambdaBufferDefinition(FnBuffer.REINVOKE_METHOD, Empty, Args, ClosureArgs)
End Function

'# Wraps a function to accept an argument array instead of a plain argument
'# This is used basically wrapped multiple arguments to one, quite hard to explain
Public Function Lambda(MethodName As Variant, _
                    Optional ClosureArgs As Variant = Empty)
    Lambda = GenerateLambdaBufferDefinition(FnBuffer.LAMBDA_METHOD, MethodName, Empty, ClosureArgs)
End Function

'# Builds the definition of the buffer
Private Function GenerateBufferDefinition(BufferMethodName As String, MethodName As Variant, BufferArgs As Variant, ClosureArgs As Variant) As String
    Dim BIndex As Long
    FnBuffer.InitializeBuffers
    BIndex = FnBuffer.GetNextBufferIndex()
    FnBuffer.SetBuffer Array( _
        BuildBufferName(BufferMethodName), MethodName, BufferArgs, ClosureArgs, BIndex), _
        BIndex
    GenerateBufferDefinition = BuildBufferName(BUFFER_PREFIX) & BIndex
End Function

'# Like GenerateBufferDefinition but for Lambda pattern
Public Function GenerateLambdaBufferDefinition(LambdaMethodName As String, MethodName As Variant, BufferArgs As Variant, ClosureArgs As Variant) As String
    Dim BIndex As Long
    FnBuffer.InitializeBuffers
    BIndex = FnBuffer.GetNextBufferIndex()
    FnBuffer.SetBuffer Array( _
        BuildLambdaBufferName(LambdaMethodName), MethodName, BufferArgs, ClosureArgs, BIndex), _
        BIndex
    GenerateLambdaBufferDefinition = BuildBufferName(BUFFER_PREFIX) & BIndex
End Function


'# Builds the full buffer module function name for use given the module and method
Private Function BuildBufferName(MethodName As String) As String
    BuildBufferName = BUFFER_MODULE & "." & MethodName
End Function
Private Function BuildLambdaBufferName(MethodName As String) As String
    BuildLambdaBufferName = LAMBDA_MODULE & "." & MethodName
End Function

