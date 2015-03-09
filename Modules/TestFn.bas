Attribute VB_Name = "TestFn"
Public Sub TestInvoke()
    VaseAssert.AssertEqual _
        Fn.Invoke("FnFunction.Identity_", Array(1)), _
        1
        
    ' Object test
    Dim Col As New Collection, PlusCol As New Collection
    Col.Add 0
    VaseAssert.AssertEqual _
        Col.Count, 1
    Set PlusCol = Fn.InvokeOneArg("FnTestLambda.AddOneToCollection_", Col)
    VaseAssert.AssertEqual _
        PlusCol.Count, 2
End Sub

Public Sub TestCurry()
    Dim AddFiveFc As Variant, AddFiveAndFourFc As Variant, AddNineFc As Variant, AddFc As Variant, ConstOneFc As Variant
    AddFiveFc = Fn.Curry("FnOperator.Add_", Array(5))
    
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFiveFc, Array(4)), _
        (5 + 4)
        
    AddFiveAndFourFc = Fn.Curry(AddFiveFc, Array(4))
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFiveAndFourFc, Array()), _
        9
        
    AddNineFc = Fn.Curry("FnOperator.Add_", Array(4, 5))
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFiveAndFourFc, Array()), _
        Fn.Invoke(AddNineFc, Array())
        
    AddFc = Fn.Curry("FnOperator.Add_", Array())
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFc, Array(4, 5)), _
        9
        
    ConstOneFc = Fn.Curry("FnFunction.Identity_", Array(1))
    VaseAssert.AssertEqual _
        Fn.Invoke(ConstOneFc, Array()), _
        1
        
    ' Object Test
    Dim Col As New Collection, PlusCol As New Collection, AddToColFc As Variant
    Col.Add 0
    VaseAssert.AssertEqual _
        Col.Count, 1
    AddToColFc = Fn.Curry("FnTestLambda.AddOneToCollection_", Array(Col))
    Set PlusCol = Fn.InvokeNoArgs(AddToColFc)
    VaseAssert.AssertEqual _
        PlusCol.Count, 2
End Sub

Public Sub TestCompose()
    Dim NegRecFc As Variant
    
    NegRecFc = Fn.Compose(Array("FnFunction.Negative_", "FnFunction.Reciprocal_"))
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(NegRecFc, 2), _
        -(1 / 2)
       
    Dim RemoveAandIFc As Variant, ToUpperAndRemoveFc As Variant
    
    RemoveAandIFc = Fn.Compose(Array("FnTestLambda.RemoveA_", "FnTestLambda.RemoveI_"))
    ToUpperAndRemoveFc = Fn.Compose(Array("FnTestLambda.ToUppercase_", RemoveAandIFc))
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(ToUpperAndRemoveFc, "Francis"), _
        "FRNCS"
    
    ' Object Test
    Dim Col As New Collection, NewCol As New Collection, DoubleAndZeroFp As Variant
    Col.Add 1
    Col.Add "Two"
    Col.Add 3#
    DoubleAndZeroFp = Fn.Compose(Array("FnTestLambda.DoubleCollection_", "FnTestLambda.AddOneToCollection_"))
    Set NewCol = Fn.InvokeOneArg(DoubleAndZeroFp, Col)
    VaseAssert.AssertEqual _
        NewCol.Count, 8
End Sub

Public Sub TestWithArgs()
    Dim WithTwoAndThreeFc As Variant
    WithTwoAndThreeFc = Fn.WithArgs(Array(2, 3))
    
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(WithTwoAndThreeFc, "FnOperator.Add_"), _
        5
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(WithTwoAndThreeFc, "FnOperator.Multiply_"), _
        6
        
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_(WithTwoAndThreeFc, _
            Array("FnOperator.Add_", "FnOperator.Multiply_")), _
        Array(5, 6)
        
    ' Object Test
    Dim Col As New Collection, NewCol As New Collection, WithMyColFc As Variant, ColArr As Variant
    Col.Add 1
    Col.Add "Two"
    Col.Add 3#
    WithMyColFc = Fn.WithArgs(Array(Col))
    ColArr = FnArrayUtil.Map_(WithMyColFc, Array("FnTestLambda.DoubleCollection_", "FnTestLambda.AddOneToCollection_"))

    VaseAssert.AssertArraySize _
        2, ColArr
    VaseAssert.AssertEqual _
        6, ColArr(0).Count
    VaseAssert.AssertEqual _
        4, ColArr(1).Count
End Sub

Public Sub TestUnpack()
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_(FnFunction.NegativeFs, ArrayUtil.Range(0, 5)), _
        ArrayUtil.Reverse(ArrayUtil.Range(-4, 1))
    
    ' Not just FnOperator.Add_Fn
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_(Fn.Unpack(FnOperator.AddFs), Array( _
            Array(1, 2), _
            Array(2, 3), _
            Array(3, 4))), _
        Array(3, 5, 7)
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_(Fn.Unpack("FnTestLambda.OperatorApply_"), Array( _
            Array(1, 2, FnOperator.AddFs), _
            Array(2, 3, FnOperator.MultiplyFs), _
            Array(3, 4, Fn.Compose(Array( _
                FnFunction.NegativeFs, FnOperator.AddFs))))), _
        Array(3, 6, -7)
        
    ' Object Test
    Dim Col As New Collection, NewCol As New Collection, WithMyColFp As Variant, ColArr As Variant
    Dim LCol As New Collection, RCol As New Collection, Arr As Variant
    LCol.Add 1
    LCol.Add 2
    RCol.Add "One"
    RCol.Add "Two"
    ColArr = FnArrayUtil.Map_(Fn.Unpack("FnTestLambda.JoinCollection_"), Array( _
                Array(LCol, RCol)))
    VaseAssert.AssertArraySize _
        1, ColArr
    VaseAssert.AssertEqual _
        ColArr(0).Count, 4
End Sub

Public Sub TestDecorate()
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg( _
            Fn.Decorate(FnFunction.NegativeFs, FnFunction.NegativeFs), _
            1), _
            1
End Sub
