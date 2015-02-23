Attribute VB_Name = "TestFnIterator"
Public Sub TestConstant()
    Dim AStream As String, EmptyStream As String
    AStream = FnIterator.Constant("A")
    EmptyStream = FnIterator.Constant(Empty)
    
    VaseAssert.AssertArraysEqual _
        FnIterator.Iterate(AStream, 5), _
        Array("A", "A", "A", "A", "A")
    
    VaseAssert.AssertArraysEqual _
        FnIterator.Iterate(EmptyStream, 5), _
        Array(Empty, Empty, Empty, Empty, Empty)
End Sub

Public Sub TestCycle()
    Dim Rng As Variant, CycleFn As String
    CycleFn = FnIterator.Cycle(Array())
    
    VaseAssert.AssertArraysEqual _
        FnIterator.Iterate(CycleFn, 3), _
        Array(Empty, Empty, Empty)
    
    Rng = ArrayUtil.Range(0, 10, 3)
    CycleFn = FnIterator.Cycle(Rng)
    
    VaseAssert.AssertArraysEqual _
        FnIterator.Iterate(CycleFn, 6), _
        Array(0, 3, 6, 9, 0, 3)
End Sub
