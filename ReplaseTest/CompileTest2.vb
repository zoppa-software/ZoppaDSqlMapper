Imports Xunit
Imports ZoppaDSqlReplace

Public Class CompileTest2

    <Fact>
    Sub ForTest()
        Dim query1 = "select {trim}{for i = 1 to 3}item!{i}, {/for}{/trim} from table1"
        Dim ans1 = query1.Replase()
        Assert.Equal("select item1, item2, item3 from table1", ans1)
    End Sub

    <Fact>
    Sub ForEachTest()
        Dim query1 = "select {trim}{foreach itm in ['A', 'B', 'C']}${itm}, {/for}{/trim} from table1"
        Dim ans1 = query1.Replase()
        Assert.Equal("select A, B, C from table1", ans1)
    End Sub

End Class
