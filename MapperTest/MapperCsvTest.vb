Imports System.Data
Imports System.Data.SQLite
Imports Xunit
Imports ZoppaDSqlreplace
Imports ZoppaDSqlreplace.Tokens
Imports ZoppaDSql
Imports ZoppaLegacyFiles.Csv

Public Class MapperCsvTest
    Implements IDisposable

    Private _sqlite As SQLiteConnection

    Public Sub New()
        _sqlite = New SQLiteConnection("Data Source=Resources\chinook.db")
        _sqlite.Open()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        _sqlite.Dispose()
    End Sub

    <Fact>
    Public Sub Case1Test()
        _sqlite.ExecuteQuery("DELETE FROM SampleDB ")

        Using tran = _sqlite.BeginTransaction()
            Dim fi As New IO.FileInfo("Resources\sample.csv")
            Using cr As New CsvStreamReader(fi.FullName)
                _sqlite.SetTransaction(tran).ExecuteQuery(Of CsvData)("INSERT INTO SampleDB (indexno, name) VALUES (@Indexno, @Name)", cr)
            End Using

            tran.Commit()
        End Using

        Dim cnt = _sqlite.ExecuteDatas(Of Long)("SELECT COUNT(*) FROM SampleDB").First()
        Assert.Equal(4, cnt)
    End Sub

    Private Class CsvData
        Public Property Indexno As Long
        Public Property Name As String
        Public Sub New(idx As Long, nm As String)
            Me.IndexNo = idx
            Me.Name = nm
        End Sub
    End Class

End Class
