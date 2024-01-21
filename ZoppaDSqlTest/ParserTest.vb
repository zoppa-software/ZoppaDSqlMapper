Imports Xunit
Imports ZoppaDSqlCompiler
Imports ZoppaDSqlCompiler.Tokens

Public Class ParserTest

    <Fact>
    Public Sub SyntaxErrorTest()
        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {end if}".Compile
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {else}".Compile
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {else if}".Compile
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {end for}".Compile
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {end trim}".Compile
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {end select}".Compile
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {case A}".Compile
            End Sub
        )
    End Sub

End Class
