Imports Microsoft.Extensions.Logging
Imports Xunit
Imports ZoppaDSqlReplace.Tokens
Imports ZoppaDSqlReplace

Public Class CalcTest2

    <Fact>
    Public Sub SignChangeTest()
        Dim op1 As New NumberToken.NumberIntegerValueToken(12)
        Dim oh_op1 = op1.SignChange()
        Assert.Equal(oh_op1.Contents, CInt(-12))

        Dim op2 As New NumberToken.NumberLongValueToken(34)
        Dim oh_op2 = op2.SignChange()
        Assert.Equal(oh_op2.Contents, CLng(-34))

        Dim op3 As New NumberToken.NumberDecimalValueToken(56.7)
        Dim oh_op3 = op3.SignChange()
        Assert.Equal(oh_op3.Contents, New Decimal(-56.7))

        Dim op4 As New NumberToken.NumberDoubleValueToken(8.9)
        Dim oh_op4 = op4.SignChange()
        Assert.Equal(oh_op4.Contents, -8.9)
    End Sub

    <Fact>
    Public Sub PlusComputationTest()
        Dim op1 As New NumberToken.NumberIntegerValueToken(12)
        Dim op1_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op1_2 As New NumberToken.NumberLongValueToken(2)
        Dim op1_3 As New NumberToken.NumberDecimalValueToken(3.45)
        Dim op1_4 As New NumberToken.NumberDoubleValueToken(6.78)
        Assert.True(op1.PlusComputation(op1_1).Contents = 13)
        Assert.True(op1.PlusComputation(op1_2).Contents = 14)
        Assert.True(op1.PlusComputation(op1_3).Contents = 15.45)
        Assert.True(op1.PlusComputation(op1_4).Contents = 18.78)

        Dim op2 As New NumberToken.NumberLongValueToken(12)
        Dim op2_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op2_2 As New NumberToken.NumberLongValueToken(2)
        Dim op2_3 As New NumberToken.NumberDecimalValueToken(3.45)
        Dim op2_4 As New NumberToken.NumberDoubleValueToken(6.78)
        Assert.True(op2.PlusComputation(op2_1).Contents = 13)
        Assert.True(op2.PlusComputation(op2_2).Contents = 14)
        Assert.True(op2.PlusComputation(op2_3).Contents = 15.45)
        Assert.True(op2.PlusComputation(op2_4).Contents = 18.78)

        Dim op3 As New NumberToken.NumberDecimalValueToken(12)
        Dim op3_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op3_2 As New NumberToken.NumberLongValueToken(2)
        Dim op3_3 As New NumberToken.NumberDecimalValueToken(3.45)
        Dim op3_4 As New NumberToken.NumberDoubleValueToken(6.78)
        Assert.True(op3.PlusComputation(op3_1).Contents = 13)
        Assert.True(op3.PlusComputation(op3_2).Contents = 14)
        Assert.True(op3.PlusComputation(op3_3).Contents = 15.45)
        Assert.True(op3.PlusComputation(op3_4).Contents = 18.78)

        Dim op4 As New NumberToken.NumberDoubleValueToken(12)
        Dim op4_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op4_2 As New NumberToken.NumberLongValueToken(2)
        Dim op4_3 As New NumberToken.NumberDecimalValueToken(3.45)
        Dim op4_4 As New NumberToken.NumberDoubleValueToken(6.78)
        Assert.True(op4.PlusComputation(op4_1).Contents = 13)
        Assert.True(op4.PlusComputation(op4_2).Contents = 14)
        Assert.True(op4.PlusComputation(op4_3).Contents = 15.45)
        Assert.True(op4.PlusComputation(op4_4).Contents = 18.78)

        Dim a105 = "+2 + +2".Executes()
        Assert.True(NumberToken.ConvertToken(4).EqualCondition(a105))

        Dim a401 = "'桃' + '太郎'".Executes().Contents
        Assert.Equal(a401.ToString(), "桃太郎")
        Dim a402 = "100 + '15'".Executes().Contents
        Assert.Equal(115, a402)
        Dim a403 = "'100' + 15".Executes().Contents
        Assert.Equal("10015", a403)
    End Sub

    <Fact>
    Public Sub MinusComputationTest()
        Dim op1 As New NumberToken.NumberIntegerValueToken(10)
        Dim op1_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op1_2 As New NumberToken.NumberLongValueToken(2)
        Dim op1_3 As New NumberToken.NumberDecimalValueToken(3.4)
        Dim op1_4 As New NumberToken.NumberDoubleValueToken(5.6)
        Assert.Equal(op1.MinusComputation(op1_1).Contents, 9)
        Assert.Equal(op1.MinusComputation(op1_2).Contents, 8)
        Assert.Equal(op1.MinusComputation(op1_3).Contents, 6.6)
        Assert.Equal(op1.MinusComputation(op1_4).Contents, 4.4)

        Dim op2 As New NumberToken.NumberLongValueToken(10)
        Dim op2_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op2_2 As New NumberToken.NumberLongValueToken(2)
        Dim op2_3 As New NumberToken.NumberDecimalValueToken(3.4)
        Dim op2_4 As New NumberToken.NumberDoubleValueToken(5.6)
        Assert.Equal(op2.MinusComputation(op2_1).Contents, 9)
        Assert.Equal(op2.MinusComputation(op2_2).Contents, 8)
        Assert.Equal(op2.MinusComputation(op2_3).Contents, 6.6)
        Assert.Equal(op2.MinusComputation(op2_4).Contents, 4.4)

        Dim op3 As New NumberToken.NumberDecimalValueToken(10)
        Dim op3_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op3_2 As New NumberToken.NumberLongValueToken(2)
        Dim op3_3 As New NumberToken.NumberDecimalValueToken(3.4)
        Dim op3_4 As New NumberToken.NumberDoubleValueToken(5.6)
        Assert.Equal(op3.MinusComputation(op3_1).Contents, 9)
        Assert.Equal(op3.MinusComputation(op3_2).Contents, 8)
        Assert.Equal(op3.MinusComputation(op3_3).Contents, 6.6)
        Assert.Equal(op3.MinusComputation(op3_4).Contents, 4.4)

        Dim op4 As New NumberToken.NumberDoubleValueToken(10)
        Dim op4_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op4_2 As New NumberToken.NumberLongValueToken(2)
        Dim op4_3 As New NumberToken.NumberDecimalValueToken(3.4)
        Dim op4_4 As New NumberToken.NumberDoubleValueToken(5.6)
        Assert.Equal(op4.MinusComputation(op4_1).Contents, 9)
        Assert.Equal(op4.MinusComputation(op4_2).Contents, 8)
        Assert.Equal(op4.MinusComputation(op4_3).Contents, 6.6)
        Assert.Equal(op4.MinusComputation(op4_4).Contents, 4.4)

        Dim a101 = "100 - 99.9".Executes()
        Assert.True(NumberToken.ConvertToken(0.1).EqualCondition(a101))
        Dim a102 = "46 - 12".Executes()
        Assert.True(NumberToken.ConvertToken(34).EqualCondition(a102))
        Dim a103 = "-2 - -2".Executes()
        Assert.True(NumberToken.ConvertToken(0).EqualCondition(a103))
        Dim a104 = "80 - '30'".Executes()
        Assert.True(NumberToken.ConvertToken(50).EqualCondition(a104))
        Dim a105 = "+2 + +2".Executes()
        Assert.True(NumberToken.ConvertToken(4).EqualCondition(a105))

        Assert.Throws(Of DSqlAnalysisException)(
            Function() "'abc' - 99".Executes()
        )
    End Sub

    <Fact>
    Public Sub MultiComputationTest()
        Dim op1 As New NumberToken.NumberIntegerValueToken(10)
        Dim op1_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op1_2 As New NumberToken.NumberLongValueToken(2)
        Dim op1_3 As New NumberToken.NumberDecimalValueToken(3.4)
        Dim op1_4 As New NumberToken.NumberDoubleValueToken(5.6)
        Assert.Equal(op1.MultiComputation(op1_1).Contents, 10)
        Assert.Equal(op1.MultiComputation(op1_2).Contents, 20)
        Assert.True(op1.MultiComputation(op1_3).Contents = 34.0)
        Assert.Equal(op1.MultiComputation(op1_4).Contents, 56.0)

        Dim op2 As New NumberToken.NumberLongValueToken(10)
        Dim op2_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op2_2 As New NumberToken.NumberLongValueToken(2)
        Dim op2_3 As New NumberToken.NumberDecimalValueToken(3.4)
        Dim op2_4 As New NumberToken.NumberDoubleValueToken(5.6)
        Assert.Equal(op2.MultiComputation(op2_1).Contents, 10)
        Assert.Equal(op2.MultiComputation(op2_2).Contents, 20)
        Assert.True(op2.MultiComputation(op2_3).Contents = 34.0)
        Assert.Equal(op2.MultiComputation(op2_4).Contents, 56.0)

        Dim op3 As New NumberToken.NumberDecimalValueToken(10)
        Dim op3_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op3_2 As New NumberToken.NumberLongValueToken(2)
        Dim op3_3 As New NumberToken.NumberDecimalValueToken(3.4)
        Dim op3_4 As New NumberToken.NumberDoubleValueToken(5.6)
        Assert.True(op3.MultiComputation(op3_1).Contents = 10)
        Assert.True(op3.MultiComputation(op3_2).Contents = 20)
        Assert.True(op3.MultiComputation(op3_3).Contents = 34.0)
        Assert.True(op3.MultiComputation(op3_4).Contents = 56.0)

        Dim op4 As New NumberToken.NumberDoubleValueToken(10)
        Dim op4_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op4_2 As New NumberToken.NumberLongValueToken(2)
        Dim op4_3 As New NumberToken.NumberDecimalValueToken(3.4)
        Dim op4_4 As New NumberToken.NumberDoubleValueToken(5.6)
        Assert.Equal(op4.MultiComputation(op4_1).Contents, 10)
        Assert.Equal(op4.MultiComputation(op4_2).Contents, 20)
        Assert.True(op4.MultiComputation(op4_3).Contents = 34.0)
        Assert.Equal(op4.MultiComputation(op4_4).Contents, 56.0)

        Dim a201 = "0.1 * 6".Executes()
        Assert.True(NumberToken.ConvertToken(0.6).EqualCondition(a201))
        Dim a202 = "26 * 2".Executes()
        Assert.True(NumberToken.ConvertToken(52).EqualCondition(a202))
        Dim a203 = "-2 * -2".Executes()
        Assert.True(NumberToken.ConvertToken(4).EqualCondition(a203))
        Dim a204 = "'2.5' * 4".Executes()
        Assert.True(NumberToken.ConvertToken(10).EqualCondition(a204))
    End Sub

    <Fact>
    Public Sub DivComputationTest()
        Dim op1 As New NumberToken.NumberIntegerValueToken(64)
        Dim op1_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op1_2 As New NumberToken.NumberLongValueToken(2)
        Dim op1_3 As New NumberToken.NumberDecimalValueToken(3.2)
        Dim op1_4 As New NumberToken.NumberDoubleValueToken(1.6)
        Assert.Equal(op1.DivComputation(op1_1).Contents, 64)
        Assert.Equal(op1.DivComputation(op1_2).Contents, 32)
        Assert.True(op1.DivComputation(op1_3).Contents = 20)
        Assert.Equal(op1.DivComputation(op1_4).Contents, 40)

        Dim op2 As New NumberToken.NumberLongValueToken(64)
        Dim op2_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op2_2 As New NumberToken.NumberLongValueToken(2)
        Dim op2_3 As New NumberToken.NumberDecimalValueToken(3.2)
        Dim op2_4 As New NumberToken.NumberDoubleValueToken(1.6)
        Assert.Equal(op2.DivComputation(op2_1).Contents, 64)
        Assert.Equal(op2.DivComputation(op2_2).Contents, 32)
        Assert.True(op2.DivComputation(op2_3).Contents = 20)
        Assert.Equal(op2.DivComputation(op2_4).Contents, 40)

        Dim op3 As New NumberToken.NumberDecimalValueToken(64)
        Dim op3_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op3_2 As New NumberToken.NumberLongValueToken(2)
        Dim op3_3 As New NumberToken.NumberDecimalValueToken(3.2)
        Dim op3_4 As New NumberToken.NumberDoubleValueToken(1.6)
        Assert.True(op3.DivComputation(op3_1).Contents = 64)
        Assert.True(op3.DivComputation(op3_2).Contents = 32)
        Assert.True(op3.DivComputation(op3_3).Contents = 20)
        Assert.True(op3.DivComputation(op3_4).Contents = 40)

        Dim op4 As New NumberToken.NumberDoubleValueToken(64)
        Dim op4_1 As New NumberToken.NumberIntegerValueToken(1)
        Dim op4_2 As New NumberToken.NumberLongValueToken(2)
        Dim op4_3 As New NumberToken.NumberDecimalValueToken(3.2)
        Dim op4_4 As New NumberToken.NumberDoubleValueToken(1.6)
        Assert.Equal(op4.DivComputation(op4_1).Contents, 64)
        Assert.Equal(op4.DivComputation(op4_2).Contents, 32)
        Assert.True(op4.DivComputation(op4_3).Contents = 20)
        Assert.Equal(op4.DivComputation(op4_4).Contents, 40)

        Dim a301 = "100 / 25".Executes()
        Assert.True(NumberToken.ConvertToken(4).EqualCondition(a301))
        Dim a302 = "'100' / 5".Executes()
        Assert.True(NumberToken.ConvertToken(20).EqualCondition(a302))
        Dim a303 = "100 / 0.5".Executes()
        Assert.True(NumberToken.ConvertToken(200).EqualCondition(a303))

        Assert.Throws(Of DivideByZeroException)(
                Function() "1 / 0".Executes()
            )
    End Sub

    <Fact>
    Sub ExpressionTest()
        Dim a301 = "100 = 100".Executes().Contents
        Assert.Equal(a301, True)
        Dim a302 = "100 = 200".Executes().Contents
        Assert.Equal(a302, False)
        Dim a303 = "'abc' = 'abc'".Executes().Contents
        Assert.Equal(a303, True)
        Dim a304 = "'abc' = 'def'".Executes().Contents
        Assert.Equal(a304, False)
        Dim a305 = "null = nothing".Executes().Contents
        Assert.Equal(a305, True)
        Dim a306 = "null = 100".Executes().Contents
        Assert.Equal(a306, False)
        Dim a307 = "'100' = 100".Executes().Contents
        Assert.Equal(a307, False)
        Dim a308 = "'100' = null".Executes().Contents
        Assert.Equal(a308, False)

        Dim a401 = "'abd' >= 'abc'".Executes().Contents
        Assert.Equal(a401, True)
        Dim a402 = "'abc' >= 'abd'".Executes().Contents
        Assert.Equal(a402, False)
        Dim a403 = "'abc' >= 'abc'".Executes().Contents
        Assert.Equal(a403, True)
        Dim a404 = "0.1+0.1+0.1 >= 0.4".Executes().Contents
        Assert.Equal(a404, False)
        Dim a405 = "0.1*5 >= 0.4".Executes().Contents
        Assert.Equal(a405, True)
        Dim a406 = "0.1*3 >= 0.3".Executes().Contents
        Assert.Equal(a406, True)

        Dim a501 = "'abd' > 'abc'".Executes().Contents
        Assert.Equal(a501, True)
        Dim a502 = "'abc' > 'abd'".Executes().Contents
        Assert.Equal(a502, False)
        Dim a503 = "'abc' > 'abc'".Executes().Contents
        Assert.Equal(a503, False)
        Dim a504 = "0.1+0.1+0.1 > 0.4".Executes().Contents
        Assert.Equal(a504, False)
        Dim a505 = "0.1*5 > 0.4".Executes().Contents
        Assert.Equal(a505, True)
        Dim a506 = "0.1*3 > 0.3".Executes().Contents
        Assert.Equal(a506, False)

        Dim a601 = "'abd' <= 'abc'".Executes().Contents
        Assert.Equal(a601, False)
        Dim a602 = "'abc' <= 'abd'".Executes().Contents
        Assert.Equal(a602, True)
        Dim a603 = "'abc' <= 'abc'".Executes().Contents
        Assert.Equal(a603, True)
        Dim a604 = "0.1+0.1+0.1 <= 0.4".Executes().Contents
        Assert.Equal(a604, True)
        Dim a605 = "0.1*5 <= 0.4".Executes().Contents
        Assert.Equal(a605, False)
        Dim a606 = "0.1*3 <= 0.3".Executes().Contents
        Assert.Equal(a606, True)

        Dim a701 = "'abd' < 'abc'".Executes().Contents
        Assert.Equal(a701, False)
        Dim a702 = "'abc' < 'abd'".Executes().Contents
        Assert.Equal(a702, True)
        Dim a703 = "'abc' < 'abc'".Executes().Contents
        Assert.Equal(a703, False)
        Dim a704 = "0.1+0.1+0.1 < 0.4".Executes().Contents
        Assert.Equal(a704, True)
        Dim a705 = "0.1*5 < 0.4".Executes().Contents
        Assert.Equal(a705, False)
        Dim a706 = "0.1*3 < 0.3".Executes().Contents
        Assert.Equal(a706, False)

        Dim a801 = "100 <> 100".Executes().Contents
        Assert.Equal(a801, False)
        Dim a802 = "100 <> 200".Executes().Contents
        Assert.Equal(a802, True)
        Dim a803 = "'abc' <> 'abc'".Executes().Contents
        Assert.Equal(a803, False)
        Dim a804 = "'abc' <> 'def'".Executes().Contents
        Assert.Equal(a804, True)
        Dim a805 = "null != nothing".Executes().Contents
        Assert.Equal(a805, False)
        Dim a806 = "null != 100".Executes().Contents
        Assert.Equal(a806, True)
        Dim a807 = "'100' != 100".Executes().Contents
        Assert.Equal(a807, True)
        Dim a808 = "'100' != null".Executes().Contents
        Assert.Equal(a808, True)
    End Sub

    <Fact>
    Sub MultOperatorTest()
        Dim a1 = "1 > 0 ? true : false".Executes().Contents
        Assert.True(a1)

        Dim a2 = "1 < 0 ? true : false".Executes().Contents
        Assert.False(a2)

        Dim a3 = "1 > 0 ? 1 : 0".Executes().Contents
        Assert.Equal(a3, 1)

        Dim a4 = "1 < 0 ? 1 : 0".Executes().Contents
        Assert.Equal(a4, 0)

        Dim a5 = "1 > 0 ? 1 + 1 : 0".Executes().Contents
        Assert.Equal(a5, 2)

        Dim a6 = "1 < 0 ? 1 + 1 : 0".Executes().Contents
        Assert.Equal(a6, 0)

        Dim a7 = "1 > 0 ? 1 + 1 : 0 + 1".Executes().Contents
        Assert.Equal(a7, 2)

        Dim a8 = "1 < 0 ? 1 + 1 : 0 + 1".Executes().Contents
        Assert.Equal(a8, 1)

        Dim a9 = "1 > 0 ? 1 + 1 : (0 + 1 ? 1 : 0)".Executes().Contents
        Assert.Equal(a9, 2)

        Assert.Throws(Of InvalidOperationException)(
            Function() "1 < 0 ? 1 + 1 : (0 + 1 ? 1 : 0)".Executes().Contents
        )

        Dim a10 = "var <> null ? var : 'null'".Executes(New With {.var = "変数"}).Contents
        Assert.Equal(a10, "変数")
    End Sub

End Class
