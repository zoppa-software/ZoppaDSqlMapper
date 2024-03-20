Imports Xunit
Imports ZoppaDSqlReplace
Imports ZoppaDSqlReplace.Tokens

Public Class ReplaseTest

    <Fact>
    Public Sub ReadmeTest()
        ' 三項
        Dim ans6 = "update person set name = #{value <> null ? value : ''}".Replase(New With {.value = Nothing})
        Assert.Equal("update person set name = ''", ans6)

        Dim ans1 = "select * from table1 where column = #{value}".Replase(New With {.value = "値"})
        Assert.Equal("select * from table1 where column = '値'", ans1)

        Dim ans2 = "select * from member where age >= #{lowAge} and age <= #{hiAge}".Replase(New With {.lowAge = 12, .hiAge = 50})
        Assert.Equal("select * from member where age >= 12 and age <= 50", ans2)

        Dim ans3 = "update person set name = #{value}".Replase(New With {.value = Nothing})
        Assert.Equal("update person set name = null", ans3)

        ' テーブル
        Dim ans4 = "select * from !{table}".Replase(New With {.table = "sample_table"})
        Assert.Equal("select * from sample_table", ans4)

        ' 条件
        Dim ans5 = "select * from table2 where !{condition}".Replase(New With {.condition = "clm1 = '123'"})
        Assert.Equal("select * from table2 where clm1 = '123'", ans5)
    End Sub

    <Fact>
    Public Sub IfTest()
        Dim query = "" &
"select * from table1
where
  {if num = 1}col1 = #{num}
  {else if num = 2}col2 = #{num}
  {else}col3 = #{num}
  {end if}"

        ' num = 1 ならば {if num = 1}の部分を出力
        Dim ans1 = query.Replase(New With {.num = 1})
        Assert.Equal(ans1,
"select * from table1
where
  col1 = 1
")

        ' num = 2 ならば {else if num = 2}の部分を出力
        Dim ans2 = query.Replase(New With {.num = 2})
        Assert.Equal(ans2,
"select * from table1
where
  col2 = 2
")

        ' num = 5 ならば {else}の部分を出力
        Dim ans3 = query.Replase(New With {.num = 5})
        Assert.Equal(ans3,
"select * from table1
where
  col3 = 5")
    End Sub

    Enum Mode
        None = 1
        MyGroup = 2
        Specified = 3
    End Enum

    <Fact>
    Public Sub SelectTest()
        Dim query = "" &
"SELECT
    *
FROM
    TBL1
{trim}
WHERE
{select mode}
{case 'None'}
{case 2}
    GRP = 0
{else}
    GRP = #{groupNo}
{/select}
{/trim}"
        Dim ans1 = query.Replase(New With {.mode = Mode.None})
        Assert.Equal(ans1,
"SELECT
    *
FROM
    TBL1
")

        Dim ans2 = query.Replase(New With {.mode = Mode.MyGroup})
        Assert.Equal(ans2,
"SELECT
    *
FROM
    TBL1
WHERE
    GRP = 0")

        Dim ans3 = query.Replase(New With {.mode = Mode.Specified, .groupNo = 100})
        Assert.Equal(ans3,
"SELECT
    *
FROM
    TBL1
WHERE
    GRP = 100")
    End Sub

End Class
