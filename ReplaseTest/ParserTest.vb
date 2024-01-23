Imports Xunit
Imports ZoppaDSqlReplace
Imports ZoppaDSqlReplace.Tokens

Public Class ParserTest

    <Fact>
    Public Sub SyntaxErrorTest()
        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {end if}".Replase
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {else}".Replase
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {else if}".Replase
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {end for}".Replase
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {end trim}".Replase
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {end select}".Replase
            End Sub
        )

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans = "select * from {case A}".Replase
            End Sub
        )
    End Sub

End Class
