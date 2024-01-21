Imports System
Imports ZoppaDSqlCompiler
Imports Microsoft.Extensions.Logging

Module Program

    Sub Main(args As String())
        Using logFactory = ZoppaDSqlCompiler.CreateZoppaDSqlLogFactory(isConsole:=True, minimumLogLevel:=LogLevel.Trace)
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
            Dim ans1 = query1.Compile(New With {.heads = {"A", "B", "C"}, .tails = {1, 2, 3, 4}})

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
".Compile(New With {.af = (i And 1) = 0, .bf = (i And 2) = 0, .cf = (i And 4) = 0, .df = (i And 8) = 0})
                Dim a As Integer = 50
            Next

            Dim a301 = "100 = 100".Executes().Contents
        End Using
    End Sub

End Module
