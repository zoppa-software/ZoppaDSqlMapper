Imports Microsoft.Extensions.Logging
Imports Xunit
Imports ZoppaDSqlReplace
Imports ZoppaDSqlReplace.Environments
Imports ZoppaDSqlReplace.Tokens

Public Class AnalysisTest2

    Public Enum SexCode
        Male = 1
        Female = 2
    End Enum

    <Fact>
    Sub EnvironmentsTest()
        Dim env As New EnvironmentObjectValue(New With {.name = "A", .number = 100, .arr = New String() {"あ", "い", "う"}})

        Assert.Equal("A", env.GetValue("name"))

        Assert.Equal(100, env.GetValue("number"))

        Assert.Equal("あ", env.GetValue("arr", 0))
        Assert.Equal("い", env.GetValue("arr", 1))
        Assert.Equal("う", env.GetValue("arr", 2))

        Assert.Throws(Of DSqlAnalysisException)(Sub() env.GetValue("arr", 3))

        Assert.Throws(Of DSqlAnalysisException)(Sub() env.GetValue("arr", -1))

        Dim env2 = env.Clone()
        env2.AddVariant("name", "B")
        Assert.Equal("A", env.GetValue("name"))
        Assert.Equal("B", env2.GetValue("name"))

        Assert.True(env2.IsDefainedName("name"))
        Assert.True(env2.IsDefainedName("number"))

        env2.AddVariant("arr", New String() {"え", "お"})
        Assert.Equal("あ", env.GetValue("arr", 0))
        Assert.Equal("い", env.GetValue("arr", 1))
        Assert.Equal("う", env.GetValue("arr", 2))
        Assert.Equal("え", env2.GetValue("arr", 0))
        Assert.Equal("お", env2.GetValue("arr", 1))

        env2.LocalVarClear()

        env2.AddVariant("name2", "C")
        Assert.Equal("C", env2.GetValue("name2"))

        env2.AddVariant("name2", "D")
        Assert.Equal("D", env2.GetValue("name2"))

        Assert.True(env2.IsDefainedName("name2"))

        env2.RemoveVariant("name2")
        Assert.False(env2.IsDefainedName("name2"))
    End Sub

    <Fact>
    Sub NotImplementedTest()
        Assert.Throws(Of NotImplementedException)(Function() AndToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() OrToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() EqualToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() NotEqualToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() GreaterEqualToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() GreaterToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() LessEqualToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() LessToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() DivToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() MultiToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() LBracketToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() LParenToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() MinusToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() NotToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() PeriodToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() PlusToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() RBracketToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() RParenToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() CommaToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() ColonToken.Value.Contents)
        Assert.Throws(Of NotImplementedException)(Function() QuestionToken.Value.Contents)
    End Sub

    <Fact>
    Sub IdentTest()
        Dim env =
                New With {
                    .group = "部署",
                    .tell = New Integer() {1, 23, 998},
                    .member = {
                        New With {.name = "A君", .sex = SexCode.Male, .age = 19},
                        New With {.name = "B君", .sex = SexCode.Male, .age = 21},
                        New With {.name = "C君", .sex = SexCode.Female, .age = 17, .ext = New With {.ch = 9}}
                    }
                }

        Dim ans1 = "group".Executes(env).Contents
        Assert.Equal(ans1, "部署")

        Dim ans2 = "tell[2]".Executes(env).Contents
        Assert.Equal(ans2, 998)

        Dim ans3 = "member[0].name".Executes(env).Contents
        Assert.Equal(ans3, "A君")

        Dim ans31 = "member[0].sex".Executes(env).Contents
        Assert.Equal(ans31, SexCode.Male)

        Dim ans4 = "member[2].ext.ch".Executes(env).Contents
        Assert.Equal(ans4, 9)
    End Sub

    <Fact>
    Sub ValueExpressTest()
        Dim a1 = "'0123'".Executes().Contents
        Assert.Equal("0123", a1)

        Dim a2 = """abcd""".Executes().Contents
        Assert.Equal("abcd", a2)
    End Sub

    <Fact>
    Sub UnaryExpressTest()
        Dim a1 = "+2".Executes()
        Assert.True(NumberToken.ConvertToken(2).EqualCondition(a1))

        Dim a2 = "-2".Executes()
        Assert.True(NumberToken.ConvertToken(-2).EqualCondition(a2))

        Dim a3 = "!false".Executes().Contents
        Assert.True(a3)

        Dim a4 = "not true".Executes().Contents
        Assert.False(a4)

        Dim a5 = "(12.3)".Executes().Contents
        Assert.Equal(12.3, a5)

        Dim a6 = "null".Executes().Contents
        Assert.Null(a6)

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "-ABC".Executes(New With {.ABC = "abc"})
                End Function
            )

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "+ABC".Executes(New With {.ABC = "abc"})
                End Function
            )

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "-null".Executes()
                End Function
            )
    End Sub

    <Fact>
    Sub ParentTest()
        Dim ans1 = "(28 - 3) / (2 + 3)".Executes().Contents
        Assert.Equal(ans1, 5)

        Dim ans2 = "(28 - 3)".Executes().Contents
        Assert.Equal(ans2, 25)
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

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "123 and '123'".Executes()
                End Function
            )

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "123 or '123'".Executes()
                End Function
            )

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "null and null".Executes()
                End Function
            )

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "100 and null".Executes()
                End Function
            )
    End Sub

    <Fact>
    Sub ArrayTest()
        Dim ans1 = "[1, 2, 3, 4]".Executes().Contents
        Assert.Equal(New Object() {1, 2, 3, 4}, ans1)

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "['あ', 'い', 'う'][1]".Executes().Contents
                End Function
            )
    End Sub

    Enum Mode
        None = 1
        MyGroup = 2
        Specified = 3
    End Enum

    <Fact>
    Sub EnumTest()
        Dim ans1 = "eval = 'None'".Executes(New With {.eval = Mode.None}).Contents
        Assert.True(ans1)

        Dim ans2 = "eval <> 'None'".Executes(New With {.eval = Mode.None}).Contents
        Assert.False(ans2)

        Dim ans3 = "eval = 3".Executes(New With {.eval = Mode.Specified}).Contents
        Assert.True(ans3)

        Dim ans4 = "eval <> 3".Executes(New With {.eval = Mode.Specified}).Contents
        Assert.False(ans4)

        Dim ans5 = "eval = eval2".Executes(New With {.eval = Mode.Specified, .eval2 = Mode.Specified}).Contents
        Assert.True(ans5)
    End Sub

    <Fact>
    Sub FormatMethodTest()
        Dim ans1 = "format('{0:#.##}', val)".Executes(New With {.val = 1.2345678}).Contents
        Assert.Equal("1.23", ans1)

        Dim ans2 = "format('123')".Executes(New With {.val = 1.2345678}).Contents
        Assert.Equal("123", ans2)

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "format()".Executes().Contents
                End Function
            )

        Assert.Throws(Of DSqlAnalysisException)(
                Function()
                    Return "unknown()".Executes().Contents
                End Function
            )
    End Sub

End Class
