Attribute VB_Name = "TestFnUtil"
Public Sub TestConstantFn()
    Dim Cn As String
    Cn = FnIterator.Constant(0)
    
    VaseAssert.AssertEqual _
        Fn.InvokeNoArgs(Cn), 0
    VaseAssert.AssertEqual _
        Fn.InvokeNoArgs(Cn), 0
End Sub
