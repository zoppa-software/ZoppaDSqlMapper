Imports Xunit
Imports ZoppaDSqlCompiler
Imports ZoppaDSqlCompiler.Environments
Imports ZoppaDSqlCompiler.Tokens

' dotnet test --collect:"XPlat Code Coverage"
' ReportGenerator -reports:"G:\ZoppaDSqlMapper\ZoppaDSqlTest\TestResults\c12e63a0-c1df-454f-8f4b-2fd0a3d7fa73\coverage.cobertura.xml" -targetdir:"coveragereport" -reporttypes:Html

Public Class AnalysisTest

    <Fact>
    Sub ParentTest()
        Dim ans1 = "(28 - 3) / (2 + 3)".Executes().Contents
        Assert.Equal(ans1, 5)

        Dim ans2 = "(28 - 3)".Executes().Contents
        Assert.Equal(ans2, 25)
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
    Sub Expression2Test()
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

        Dim a201 = "0.1 * 6".Executes()
        Assert.True(NumberToken.ConvertToken(0.6).EqualCondition(a201))
        Dim a202 = "26 * 2".Executes()
        Assert.True(NumberToken.ConvertToken(52).EqualCondition(a202))
        Dim a203 = "-2 * -2".Executes()
        Assert.True(NumberToken.ConvertToken(4).EqualCondition(a203))
        Dim a204 = "'2.5' * 4".Executes()
        Assert.True(NumberToken.ConvertToken(10).EqualCondition(a204))

        Dim a301 = "100 / 25".Executes()
        Assert.True(NumberToken.ConvertToken(4).EqualCondition(a301))
        Dim a302 = "'100' / 5".Executes()
        Assert.True(NumberToken.ConvertToken(20).EqualCondition(a302))

        Dim a401 = "'桃' + '太郎'".Executes().Contents
        Assert.Equal(a401, "桃太郎")
        Dim a402 = "100 + '15'".Executes().Contents
        Assert.Equal(115, a402)

        Dim a501 = "#2024/1/1# > dateValue".Executes(New With {.dateValue = New Date(2023, 1, 1)}).Contents
        Assert.True(a501)
    End Sub

    <Fact>
    Sub BinaryTest()
        Dim a101 = "true and true".Executes().Contents
        Assert.Equal(a101, True)
        Dim a102 = "true and false".Executes().Contents
        Assert.Equal(a102, False)
        Dim a103 = "false and true".Executes().Contents
        Assert.Equal(a103, False)
        Dim a104 = "false and false".Executes().Contents
        Assert.Equal(a104, False)
        Dim a105 = "false and !false".Executes().Contents
        Assert.Equal(a105, False)

        Dim a201 = "true or true".Executes().Contents
        Assert.Equal(a201, True)
        Dim a202 = "true or false".Executes().Contents
        Assert.Equal(a202, True)
        Dim a203 = "false or true".Executes().Contents
        Assert.Equal(a203, True)
        Dim a204 = "false or false".Executes().Contents
        Assert.Equal(a204, False)
        Dim a205 = "false or !false".Executes().Contents
        Assert.Equal(a205, True)
    End Sub

    <Fact>
    Sub ExceptionTest()
        Try
            Dim a0 = "1 / 0".Executes()
            Assert.True(False, "0割例外が発生しない")
        Catch ex As DivideByZeroException

        End Try

        Try
            Dim a1 = "123 and '123'".Executes()
            Assert.True(False, "論理積エラーが発生しない")
        Catch ex As DSqlAnalysisException

        End Try

        Try
            Dim a1 = "123 or '123'".Executes()
            Assert.True(False, "論理和エラーが発生しない")
        Catch ex As DSqlAnalysisException

        End Try

        Try
            Dim a3 = "'abc' - 99".Executes()
            Assert.True(False, "減算エラーが発生しない")
        Catch ex As DSqlAnalysisException

        End Try

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim a = "null and null".Executes()
            End Sub
        )
        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim a = "100 and null".Executes()
            End Sub
        )
    End Sub

    <Fact>
    Sub EnvironmentTest()
        Dim env As New EnvironmentObjectValue(New With {.prm1 = 100, .prm2 = "abc"})
        Assert.Equal(env.GetValue("prm1"), 100)
        Assert.Equal(env.GetValue("prm2"), "abc")

        Assert.Equal(env.GetValue("prm2"), "abc")

        env.AddVariant("prm3", "def")
        Assert.Equal(env.GetValue("prm3"), "def")

        env.AddVariant("prm3", "xyz")
        Assert.Equal(env.GetValue("prm3"), "xyz")

        env.LocalVarClear()
        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Assert.Equal(env.GetValue("prm3"), "def")
            End Sub
        )
    End Sub

    <Fact>
    Sub EnvironmentTest2()
        Dim a1 = "prm + 5".Executes(New With {.prm = 60}).Contents
        Assert.Equal(65, a1)

        Dim a2 = "prm - 10".Executes(New With {.prm = 60}).Contents
        Assert.Equal(50, a2)

        Dim a3 = "prm * 10".Executes(New With {.prm = 6}).Contents
        Assert.Equal(60, a3)

        Dim a4 = "prm / 5".Executes(New With {.prm = 60}).Contents
        Assert.Equal(12, a4)

        Dim a11 = "5 + prm".Executes(New With {.prm = 60}).Contents
        Assert.Equal(65, a11)

        Dim a12 = "100 - prm".Executes(New With {.prm = 60}).Contents
        Assert.Equal(40, a12)

        Dim a13 = "10 * prm".Executes(New With {.prm = 6}).Contents
        Assert.Equal(60, a13)

        Dim a14 = "60 / prm".Executes(New With {.prm = 5}).Contents
        Assert.Equal(12, a14)
    End Sub

End Class
