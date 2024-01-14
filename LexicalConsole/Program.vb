Imports System
Imports ZoppaDSqlCompiler
Imports Microsoft.Extensions.Logging

Module Program

    Sub Main(args As String())
        Using logFactory = ZoppaDSqlCompiler.CreateZoppaDSqlLogFactory(isConsole:=True, minimumLogLevel:=LogLevel.Trace)
            Dim a = "#2024/1/1# > dateValue".Executes(New With {.dateValue = New Date(2023, 1, 1)}).Contents

            Dim a301 = "100 = 100".Executes().Contents
        End Using
    End Sub

End Module
