Attribute VB_Name = "TestFn"
Public Sub TestInvoke()
    VaseAssert.AssertEqual _
        Fn.Invoke("FnFunction.Identity_", Array(1)), _
        1
End Sub

Public Sub TestCurry()
    Dim AddFive As String, AddFiveAndFour As String, AddNine As String, Add_L As String
    Dim ConstOne As String
    AddFive = Fn.Curry("FnOperator.Add_", Array(5))
    
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFive, Array(4)), _
        (5 + 4)
        
    AddFiveAndFour = Fn.Curry(AddFive, Array(4))
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFiveAndFour, Array()), _
        9
        
    AddNine = Fn.Curry("FnOperator.Add_", Array(4, 5))
    VaseAssert.AssertEqual _
        Fn.Invoke(AddFiveAndFour, Array()), _
        Fn.Invoke(AddNine, Array())
        
    Add_L = Fn.Curry("FnOperator.Add_", Array())
    VaseAssert.AssertEqual _
        Fn.Invoke(Add_L, Array(4, 5)), _
        9
        
    ConstOne = Fn.Curry("FnFunction.Identity_", Array(1))
    VaseAssert.AssertEqual _
        Fn.Invoke(ConstOne, Array()), _
        1
End Sub

Public Sub TestCompose()
    Dim NegRecFn As String
    
    NegRecFn = Fn.Compose(Array("FnFunction.Negative_", "FnFunction.Reciprocal_"))
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(NegRecFn, 2), _
        -(1 / 2)
       
    Dim RemoveAandI_Fn As String, ToUpperAndRemove_Fn As String
    
    RemoveAandI_Fn = Fn.Compose(Array("FnTestLambda.RemoveA_", "FnTestLambda.RemoveI_"))
    ToUpperAndRemove_Fn = Fn.Compose(Array("FnTestLambda.ToUppercase_", RemoveAandI_Fn))
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(ToUpperAndRemove_Fn, "Francis"), _
        "FRNCS"
End Sub

Public Sub TestReinvoke()
    Dim WithTwoAndThree As String
    WithTwoAndThree = Fn.Reinvoke(Array(2, 3))
    
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(WithTwoAndThree, "FnOperator.Add_"), _
        5
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg(WithTwoAndThree, "FnOperator.Multiply_"), _
        6
        
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_(WithTwoAndThree, _
            Array("FnOperator.Add_", "FnOperator.Multiply_")), _
        Array(5, 6)
End Sub

Public Sub TestLambda()
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_(FnFunction.Negative_Fn, ArrayUtil.Range(0, 5)), _
        ArrayUtil.Reverse(ArrayUtil.Range(-4, 1))
    
    ' Not just FnOperator.Add_Fn
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_(Fn.Lambda(FnOperator.Add_Fn), Array( _
            Array(1, 2), _
            Array(2, 3), _
            Array(3, 4))), _
        Array(3, 5, 7)
    VaseAssert.AssertArraysEqual _
        FnArrayUtil.Map_(Fn.Lambda("FnTestLambda.OperatorApply_"), Array( _
            Array(1, 2, FnOperator.Add_Fn), _
            Array(2, 3, FnOperator.Multiply_Fn), _
            Array(3, 4, Fn.Compose(Array( _
                FnFunction.Negative_Fn, FnOperator.Add_Fn))))), _
        Array(3, 6, -7)
End Sub

Public Sub TestDecorate()
    VaseAssert.AssertEqual _
        Fn.InvokeOneArg( _
            Fn.Decorate(FnFunction.Negative_Fn, FnFunction.Negative_Fn), _
            1), _
            1
End Sub
