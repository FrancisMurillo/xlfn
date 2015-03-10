Attribute VB_Name = "TestFnUtil"
Public Sub TestMemoizeAndTimeIt()
    Dim MemoFibonacciFp As Variant
    MemoFibonacciFp = FnUtil.Timeit(FnUtil.Memoize("TestFnUtil.Fibonacci_"))
    
    Dim FirstRes As Variant, FirstVal As Variant, SecondRes As Variant
    FirstVal = Fibonacci(20)
    FirstRes = Fn.InvokeOneArg(MemoFibonacciFp, 20)
    SecondRes = Fn.InvokeOneArg(MemoFibonacciFp, 20)
    
    ' Should return equal value
    VaseAssert.AssertEqual _
        FirstVal, FirstRes(0)
        
    ' Should be still the same
    VaseAssert.AssertEqual _
        FirstRes(0), SecondRes(0)
    ' Should be faster
    VaseAssert.AssertGreaterThan _
        FirstRes(1), SecondRes(1)
        
    Ping_
End Sub
Public Function Fibonacci(N As Long)
    If N < 3 Then
        Fibonacci = 1
    Else
        Fibonacci = Fibonacci(N - 1) + Fibonacci(N - 2)
    End If
End Function
Public Sub Fibonacci_(N As Long)
    If N < 3 Then
        Fn.Result = 1
    Else
        Fn.Result = Fn.InvokeOneArg(Fn.ThisFp, N - 1) + Fn.InvokeOneArg(Fn.ThisFp, N - 2)
    End If
End Sub
