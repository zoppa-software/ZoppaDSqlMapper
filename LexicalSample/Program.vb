Imports System
Imports ZoppaDSqlReplace
Imports Microsoft.Extensions.Logging

Module Program

    Sub Main(args As String())
        Using logFactory = ZoppaDSqlReplace.CreateZoppaDSqlLogFactory(isConsole:=True, minimumLogLevel:=LogLevel.Trace)
            Dim query4 = "" &
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
            Dim ans4 = query4.Replase(New With {.mode = 2})

            Dim query1 = "" &
"SELECT
    {trim}
    {for each head in heads}
        {for each tail in tails}
            ${head}_${tail},
        {end for}
    {end for}
    {end trim}
FROM
    TEST_TABLE"
            Dim ans1 = query1.Replase(New With {.heads = {"A", "B", "C"}, .tails = {1, 2, 3, 4}})

            For i As Integer = 0 To 15
                Dim ans2 = "" &
"select * from tb1
{trim 'where', 'in'}
where
    {trim both}
        {trim}
            ({trim both}{if af}a = 1{/if} or {if bf}b = 2{/if}{/trim})
        {/trim}
        and
        {trim}
            ({trim both}{if cf}c = 3{/if} or {if df}d = 4{/if}{/trim})
        {/trim}
    {/trim}
{/trim}
".Replase(New With {.af = (i And 1) = 0, .bf = (i And 2) = 0, .cf = (i And 4) = 0, .df = (i And 8) = 0})
                ' ‹óŽÀ‘•
            Next

            Dim a301 = "100 = 100".Executes().Contents

        End Using
    End Sub

End Module
