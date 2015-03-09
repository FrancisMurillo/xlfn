Attribute VB_Name = "TestFnIterator"
Public Sub TestConstant()
    Dim AStream As Variant, EmptyStream As Variant, ObjectStream As Variant
    AStream = FnIterator.Constant("A")
    EmptyStream = FnIterator.Constant(Empty)
    ObjectStream = FnIterator.Constant(New Collection)
         
    VaseAssert.AssertArraysEqual _
        FnIterator.Iterate(AStream, 5), _
        Array("A", "A", "A", "A", "A")
    
    VaseAssert.AssertArraysEqual _
        FnIterator.Iterate(EmptyStream, 5), _
        Array(Empty, Empty, Empty, Empty, Empty)
    
    VaseAssert.AssertArraySize _
        5, FnIterator.Iterate(ObjectStream, 5)

End Sub

Public Sub TestRandom()
    Dim RStream As Variant, RVal As Variant
    RStream = FnIterator.Iterate(FnIterator.Random(0, 10), 10)
    
    For Each RVal In RStream
        VaseAssert.AssertLessThanOrEqual RVal, 10
        VaseAssert.AssertGreaterThanOrEqual RVal, 0
    Next
End Sub

Public Sub TestCycle()
    Dim Rng As Variant, CycleFp As Variant
    CycleFp = FnIterator.Cycle(Array())
    
    VaseAssert.AssertArraysEqual _
        FnIterator.Iterate(CycleFp, 3), _
        Array(Empty, Empty, Empty)
    
    Rng = ArrayUtil.Range(0, 10, 3)
    CycleFp = FnIterator.Cycle(Rng)
     
    VaseAssert.AssertArraysEqual _
        FnIterator.Iterate(CycleFp, 6), _
        Array(0, 3, 6, 9, 0, 3)
End Sub
