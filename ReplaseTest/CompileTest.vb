﻿Imports Xunit
Imports ZoppaDSqlReplace
Imports ZoppaDSqlReplace.Tokens

Public Class CompileTest

    ' 改行コードが CRLFからLFに変わっているかもしれません。その場合はCRLFに変更するとテストが通ります

    <Fact>
    Public Sub WhereTrimTest()
        Dim query3 = "" &
"select * from tb1
{trim}
where
    {trim both}
        {trim}({trim both}{if af}a = 1{end if} or {if bf}b = 2{end if}{/trim}){/trim} and {trim}({trim both}{if cf}c = 3{end if} or {if df}d = 4{end if}{/trim}){/trim}
    {/trim}
{end trim}
"
        Dim ans7 = query3.Replase(New With {.af = True, .bf = False, .cf = False, .df = True})
        Assert.Equal(ans7.Trim(),
"select * from tb1
where
        (a = 1) and (d = 4)")

        Dim ans6 = query3.Replase(New With {.af = False, .bf = False, .cf = False, .df = False})
        Assert.Equal(ans6.Trim(), "select * from tb1")

        Dim query1 = "" &
"select * from employees
{trim}
where
    {trim both}
    {if empNo}emp_no < 20000{end if} and
    {trim}
      ({trim both}{if first_name}first_name like 'A%'{end if} or {if gender}gender = 'F'{end if}{/trim})
    {/trim}
    {/trim}
{end trim}
limit 10"
        Dim ans1 = query1.Replase(New With {.empNo = False, .first_name = False, .gender = False})
        Assert.Equal(ans1.Trim(),
"select * from employees
limit 10")

        Dim ans2 = query1.Replase(New With {.empNo = True, .first_name = False, .gender = False})
        Assert.Equal(ans2.Trim(),
"select * from employees
where
    emp_no < 20000
limit 10")

        Dim ans3 = query1.Replase(New With {.empNo = False, .first_name = True, .gender = False})
        Assert.Equal(ans3.Trim(),
"select * from employees
where
      (first_name like 'A%')
limit 10")

        Dim query2 = "" &
"select * from employees
{trim}
where
    {trim}({if empNo}emp_no = sysdate(){end if}){/trim}
{end trim}
limit 10"
        Dim ans4 = query2.Replase(New With {.empNo = False})
        Assert.Equal(ans4.Trim(),
"select * from employees
limit 10")

        Dim ans5 = query2.Replase(New With {.empNo = True})
        Assert.Equal(ans5.Trim(),
"select * from employees
where
    (emp_no = sysdate())
limit 10")
    End Sub

    <Fact>
    Public Sub ConnmaTrimTest()
        Dim query1 = "" &
"SELECT
    *
FROM
    customers
WHERE
    FirstName in ({trim}{foreach nm in names}#{nm}{}, {end for}{end trim})
"
        Dim ans1 = query1.Replase(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
        Assert.Equal(ans1,
"SELECT
    *
FROM
    customers
WHERE
    FirstName in ('Helena', 'Dan', 'Aaron')
")

        Dim query2 = "" &
"SELECT
    *
FROM
    customers
WHERE
    FirstName in (#{names})
"
        Dim ans2 = query2.Replase(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
        Assert.Equal(ans2.Trim(),
"SELECT
    *
FROM
    customers
WHERE
    FirstName in ('Helena', 'Dan', 'Aaron')")
    End Sub

    <Fact>
    Public Sub CompileErrorTest()
        Dim query1 = "" &
"SELECT
    *
FROM
    customers 
WHERE
    FirstName in ({trim}
        {foreach nm in names}#{nm}{}, 
        {end foe}
    {end trim})
"
        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                query1.Replase(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
            End Sub
        )
    End Sub

    <Fact>
    Public Sub ReplaseTest()
        Dim ansstr1 = "where parson_name = #{name}".Replase(New With {.name = "zouta takshi"})
        Assert.Equal(ansstr1, "where parson_name = 'zouta takshi'")

        Dim ansnum1 = "where age >= #{age}".Replase(New With {.age = 12})
        Assert.Equal(ansnum1, "where age >= 12")

        Dim ansnull = "set col1 = #{null}".Replase(New With {.null = Nothing})
        Assert.Equal(ansnull, "set col1 = null")

        Dim ansstr2 = "select * from !{table}".Replase(New With {.table = "sample_table"})
        Assert.Equal(ansstr2, "select * from sample_table")

        Dim ansstr3 = "select * from sample_${table}".Replase(New With {.table = "table"})
        Assert.Equal(ansstr3, "select * from sample_table")

        Dim ans1 = "a = #{'123'}".Replase()
        Assert.Equal(ans1, "a = '123'")

        Dim ans2 = "b = #{9 * 9}".Replase()
        Assert.Equal(ans2, "b = 81")

        Dim ans3 = "b = #{9 * num}".Replase(New With {.num = 3})
        Assert.Equal(ans3, "b = 27")
    End Sub

    <Fact>
    Public Sub IfTest()
        Dim query = "" &
"select * from table1
where
  {if num = 1}col1 = #{num}
  {else if num = 2}col2 = #{num}
  {else}col3 = #{num}
  {end if}
"
        Dim ans1 = query.Replase(New With {.num = 1})
        Assert.Equal(ans1,
"select * from table1
where
  col1 = 1
")
        Dim ans2 = query.Replase(New With {.num = 2})
        Assert.Equal(ans2,
"select * from table1
where
  col2 = 2
")
        Dim ans3 = query.Replase(New With {.num = 5})
        Assert.Equal(ans3,
"select * from table1
where
  col3 = 5
")
    End Sub

    <Fact>
    Public Sub SelectTest()
        Dim query = "" &
"SELECT
    *
FROM
{select mode}
    {case 1}
    TEST_TABLE1
    {case 2}
    TEST_TABLE2
    {else}
    TEST_TABLE0
{/select}"
        Dim ans1 = query.Replase(New With {.mode = 1})
        Assert.Equal(ans1.Trim(),
"SELECT
    *
FROM
    TEST_TABLE1")

        Dim ans2 = query.Replase(New With {.mode = 2})
        Assert.Equal(ans2.Trim(),
"SELECT
    *
FROM
    TEST_TABLE2")

        Dim ans3 = query.Replase(New With {.mode = 3})
        Assert.Equal(ans3.Trim(),
"SELECT
    *
FROM
    TEST_TABLE0")
    End Sub

    <Fact>
    Public Sub TrimTest()
        Dim ans1 = "{trim ""tr""}   a = #{12 * 13}   {end trim}".Replase()
        Assert.Equal(ans1, "   a = 156   ")

        Dim ans2 = "{trim 'trush'}   a = #{'11' + '29'}trush{end trim}".Replase()
        Assert.Equal(ans2, "   a = '1129'")

        Dim strs As New List(Of String)()
        strs.Add("あいうえお")
        strs.Add("かきくけこ")
        strs.Add("さしすせそ")
        strs.Add("たちつてと")
        strs.Add("なにぬねの")
        Dim ans3 = "{trim}
{foreach str in strs}
    #{str},
{end for}
{end trim}".Replase(New With {.strs = strs})
        Assert.Equal(ans3,
"    'あいうえお',
    'かきくけこ',
    'さしすせそ',
    'たちつてと',
    'なにぬねの'")

        Dim ans4 = "{trim 'tr',123}a = 100123{end trim}".Replase()
        Assert.Equal(ans4, "a = 100")

        Assert.Throws(Of DSqlAnalysisException)(
            Sub()
                Dim ans5 = "{trim 'tr',123, ABC}a = 100123{end trim}".Replase()
            End Sub
        )
    End Sub

    <Fact>
    Public Sub TrimTest2()
        Dim query = "SELECT
{if sel}
	Tekiyo.Code,Tekiyo.Name
{else}
	count(*) as cnt
{end if}
FROM Tekiyo
{trim}
WHERE
    {trim both}
    {if txZyCd <> ''}Tekiyo.担当者CD like '%!{txZyCd}%'{end if}
    {if txZyNm <> ''}AND Tekiyo.Name like '%!{txZyNm}%'{end if}
    {/trim}
{end trim}
{if sel}
ORDER BY Tekiyo.Code
{end if}
"
        Dim ans1 = query.Replase(New With {.sel = True, .txZyCd = "A", .txZyNm = ""})
        Assert.Equal(ans1, "SELECT
	Tekiyo.Code,Tekiyo.Name
FROM Tekiyo
WHERE
    Tekiyo.担当者CD like '%A%'
ORDER BY Tekiyo.Code
")

        Dim ans2 = query.Replase(New With {.sel = True, .txZyCd = "", .txZyNm = ""})
        Assert.Equal(ans2, "SELECT
	Tekiyo.Code,Tekiyo.Name
FROM Tekiyo
ORDER BY Tekiyo.Code
")

        Dim ans3 = query.Replase(New With {.sel = True, .txZyCd = "", .txZyNm = "B"})
        Assert.Equal(ans3, "SELECT
	Tekiyo.Code,Tekiyo.Name
FROM Tekiyo
WHERE
     Tekiyo.Name like '%B%'
ORDER BY Tekiyo.Code
")

        Dim ans4 = query.Replase(New With {.sel = True, .txZyCd = "A", .txZyNm = "B"})
        Assert.Equal(ans4, "SELECT
	Tekiyo.Code,Tekiyo.Name
FROM Tekiyo
WHERE
    Tekiyo.担当者CD like '%A%'
    AND Tekiyo.Name like '%B%'
ORDER BY Tekiyo.Code
")
    End Sub

    <Fact>
    Public Sub TrimTest3()
        Dim query = "SELECT
{if sel}
	Tekiyo.Code,Tekiyo.Name,Tekiyo.No
{else}
	count(*) as cnt
{end if}
FROM Tekiyo
{trim}
WHERE
    {trim both}
    {if txZyCd <> ''}Tekiyo.担当者CD like '%!{txZyCd}%'{end if}
    {if txZyNm <> ''}AND Tekiyo.Name like '%!{txZyNm}%'{end if}
    {if txZyNo <> ''}AND Tekiyo.No like '%!{txZyNo}%'{end if}
    {/trim}
{end trim}
{if sel}
ORDER BY Tekiyo.Code,Tekiyo.No
{end if}
"
        Dim ans1 = query.Replase(New With {.sel = True, .txZyCd = "A", .txZyNm = "", .txZyNo = ""})
        Assert.Equal(ans1.Trim(), "SELECT
	Tekiyo.Code,Tekiyo.Name,Tekiyo.No
FROM Tekiyo
WHERE
    Tekiyo.担当者CD like '%A%'
ORDER BY Tekiyo.Code,Tekiyo.No")

        Dim ans2 = query.Replase(New With {.sel = True, .txZyCd = "", .txZyNm = "B", .txZyNo = ""})
        Assert.Equal(ans2.Trim(), "SELECT
	Tekiyo.Code,Tekiyo.Name,Tekiyo.No
FROM Tekiyo
WHERE
     Tekiyo.Name like '%B%'
ORDER BY Tekiyo.Code,Tekiyo.No")

        Dim ans3 = query.Replase(New With {.sel = True, .txZyCd = "", .txZyNm = "", .txZyNo = "C"})
        Assert.Equal(ans3.Trim(), "SELECT
	Tekiyo.Code,Tekiyo.Name,Tekiyo.No
FROM Tekiyo
WHERE
     Tekiyo.No like '%C%'
ORDER BY Tekiyo.Code,Tekiyo.No")

        Dim ans4 = query.Replase(New With {.sel = True, .txZyCd = "A", .txZyNm = "B", .txZyNo = ""})
        Assert.Equal(ans4.Trim(), "SELECT
	Tekiyo.Code,Tekiyo.Name,Tekiyo.No
FROM Tekiyo
WHERE
    Tekiyo.担当者CD like '%A%'
    AND Tekiyo.Name like '%B%'
ORDER BY Tekiyo.Code,Tekiyo.No")

        Dim ans5 = query.Replase(New With {.sel = True, .txZyCd = "A", .txZyNm = "", .txZyNo = "C"})
        Assert.Equal(ans5.Trim(), "SELECT
	Tekiyo.Code,Tekiyo.Name,Tekiyo.No
FROM Tekiyo
WHERE
    Tekiyo.担当者CD like '%A%'
    AND Tekiyo.No like '%C%'
ORDER BY Tekiyo.Code,Tekiyo.No")

        Dim ans6 = query.Replase(New With {.sel = True, .txZyCd = "", .txZyNm = "B", .txZyNo = "C"})
        Assert.Equal(ans6.Trim(), "SELECT
	Tekiyo.Code,Tekiyo.Name,Tekiyo.No
FROM Tekiyo
WHERE
     Tekiyo.Name like '%B%'
    AND Tekiyo.No like '%C%'
ORDER BY Tekiyo.Code,Tekiyo.No")

        Dim ans7 = query.Replase(New With {.sel = True, .txZyCd = "A", .txZyNm = "B", .txZyNo = "C"})
        Assert.Equal(ans7.Trim(), "SELECT
	Tekiyo.Code,Tekiyo.Name,Tekiyo.No
FROM Tekiyo
WHERE
    Tekiyo.担当者CD like '%A%'
    AND Tekiyo.Name like '%B%'
    AND Tekiyo.No like '%C%'
ORDER BY Tekiyo.Code,Tekiyo.No")

        Dim ans8 = query.Replase(New With {.sel = True, .txZyCd = "", .txZyNm = "", .txZyNo = ""})
        Assert.Equal(ans8.Trim(), "SELECT
	Tekiyo.Code,Tekiyo.Name,Tekiyo.No
FROM Tekiyo
ORDER BY Tekiyo.Code,Tekiyo.No")
    End Sub

End Class
