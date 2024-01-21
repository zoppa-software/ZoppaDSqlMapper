Imports System.Data
Imports System.Data.SQLite
Imports Xunit
Imports ZoppaDSql

Public Class SQLTest
    Implements IDisposable

    Private mSQLite As SQLiteConnection

    Public Sub New()
        Me.mSQLite = New SQLiteConnection("Data Source=sample.db")
        Me.mSQLite.Open()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Me.mSQLite?.Dispose()
    End Sub

    <Fact>
    Public Sub CUIDTest()
        Dim tran1 = Me.mSQLite.BeginTransaction()
        Try
            Me.mSQLite.SetTransaction(tran1).ExecuteQuery(
"CREATE TABLE Zodiac (
name TEXT,
jp_name TEXT,
from_date DATETIME NOT NULL,
to_date DATETIME NOT NULL,
PRIMARY KEY(name)
)")
            tran1.Commit()
        Catch ex As Exception
            tran1.Rollback()
        End Try

        Dim zodiacs = New Zodiac() {
            New Zodiac("Aries", "牡羊座", New Date(2022, 3, 21), New Date(2022, 4, 19)),
            New Zodiac("Taurus", "牡牛座", New Date(2022, 4, 20), New Date(2022, 5, 20)),
            New Zodiac("Gemini", "双子座", New Date(2022, 5, 21), New Date(2022, 6, 21)),
            New Zodiac("Cancer", "蟹座", New Date(2022, 6, 22), New Date(2022, 7, 22)),
            New Zodiac("Leo", "獅子座", New Date(2022, 7, 23), New Date(2022, 8, 22)),
            New Zodiac("Virgo", "乙女座", New Date(2022, 8, 23), New Date(2022, 9, 22)),
            New Zodiac("Libra", "天秤座", New Date(2022, 9, 23), New Date(2022, 10, 23)),
            New Zodiac("Scorpio", "蠍座", New Date(2022, 10, 24), New Date(2022, 11, 22)),
            New Zodiac("Sagittarius", "射手座", New Date(2022, 11, 23), New Date(2022, 12, 21)),
            New Zodiac("Capricom", "山羊座", New Date(2022, 12, 22), New Date(2023, 1, 19)),
            New Zodiac("Aquuarius", "水瓶座", New Date(2023, 1, 20), New Date(2023, 2, 18)),
            New Zodiac("Pisces", "魚座", New Date(2023, 2, 19), New Date(2023, 3, 20))
        }

        Dim tran = Me.mSQLite.BeginTransaction()
        Try
            Me.mSQLite.SetTransaction(tran).ExecuteQuery(
                "INSERT INTO Zodiac (name, jp_name, from_date, to_date) 
                    VALUES (@Name, @JpName, @FromDate, @ToDate)", Nothing, zodiacs)
            tran.Commit()
        Catch ex As Exception
            tran.Rollback()
        End Try

        Dim ansZodiacs = Me.mSQLite.ExecuteRecords(Of Zodiac)(
            "select name, jp_name, from_date, to_date from Zodiac"
        )
        Assert.Equal(zodiacs, ansZodiacs.ToArray())

        Dim ansCount = Me.mSQLite.ExecuteDatas(Of Long)(
            "select count(*) from Zodiac"
        )
        Assert.Equal(12, ansCount.First())

        Dim ansObjects = Me.mSQLite.ExecuteArrays(
            "select name, jp_name, from_date, to_date from Zodiac"
        )
        Assert.Equal(12, ansObjects.Count)

        Dim selZodiacs = Me.mSQLite.ExecuteRecords(Of Zodiac)(
            "select name, jp_name, from_date, to_date from Zodiac where jp_name = @jname",
            New With {.jname = CType("双子座", DbString)}
        )
        Assert.Equal(New Zodiac() {zodiacs(2)}, selZodiacs.ToArray())

        Dim tran3 = Me.mSQLite.BeginTransaction()
        Try
            Me.mSQLite.SetTransaction(tran3).ExecuteQuery(
"DROP TABLE Zodiac")
            tran3.Commit()
        Catch ex As Exception
            tran3.Rollback()
        End Try
    End Sub

    <Fact>
    Public Sub KeyTest()
        Dim key1 = UniqueKey.Create(1, "ABC", New Date(2022, 9, 24))
        Dim key2 = UniqueKey.Create(1, Nothing, New Date(2022, 9, 24))
        Dim comp1 = key1.CompareTo(key2)
        Assert.Equal(comp1, 1)
    End Sub

    <Fact>
    Public Sub CreateMethodTest()
        Try
            Me.mSQLite.ExecuteQuery("DROP TABLE Zodiac")
            Me.mSQLite.ExecuteQuery("DROP TABLE Person")
        Catch ex As Exception

        End Try

        Dim tran1 = Me.mSQLite.BeginTransaction()
        Try
            Me.mSQLite.SetTransaction(tran1).ExecuteQuery(
"CREATE TABLE Zodiac (
name TEXT,
jp_name TEXT,
from_date DATETIME NOT NULL,
to_date DATETIME NOT NULL,
PRIMARY KEY(name)
)")
            Me.mSQLite.SetTransaction(tran1).ExecuteQuery(
"CREATE TABLE Person (
name TEXT,
zodiac TEXT,
birth_day DATETIME NOT NULL,
PRIMARY KEY(name)
)")

            tran1.Commit()
        Catch ex As Exception
            tran1.Rollback()
        End Try

        Dim zodiacs = New Zodiac() {
            New Zodiac("Aries", "牡羊座", New Date(2022, 3, 21), New Date(2022, 4, 19)),
            New Zodiac("Taurus", "牡牛座", New Date(2022, 4, 20), New Date(2022, 5, 20)),
            New Zodiac("Gemini", "双子座", New Date(2022, 5, 21), New Date(2022, 6, 21)),
            New Zodiac("Cancer", "蟹座", New Date(2022, 6, 22), New Date(2022, 7, 22)),
            New Zodiac("Leo", "獅子座", New Date(2022, 7, 23), New Date(2022, 8, 22)),
            New Zodiac("Virgo", "乙女座", New Date(2022, 8, 23), New Date(2022, 9, 22)),
            New Zodiac("Libra", "天秤座", New Date(2022, 9, 23), New Date(2022, 10, 23)),
            New Zodiac("Scorpio", "蠍座", New Date(2022, 10, 24), New Date(2022, 11, 22)),
            New Zodiac("Sagittarius", "射手座", New Date(2022, 11, 23), New Date(2022, 12, 21)),
            New Zodiac("Capricom", "山羊座", New Date(2022, 12, 22), New Date(2023, 1, 19)),
            New Zodiac("Aquuarius", "水瓶座", New Date(2023, 1, 20), New Date(2023, 2, 18)),
            New Zodiac("Pisces", "魚座", New Date(2023, 2, 19), New Date(2023, 3, 20))
        }

        Dim persons = New Person() {
            New Person("佐藤 健", "Aries", New Date(1989, 3, 21)),
            New Person("岩城 滉一", "Aries", New Date(1951, 3, 21)),
            New Person("大橋 巨泉", "Aries", New Date(1934, 3, 22)),
            New Person("阿部 サダヲ", "Taurus", New Date(1970, 4, 23)),
            New Person("大沢 樹生", "Taurus", New Date(1969, 4, 20)),
            New Person("有吉 弘行", "Gemini", New Date(1974, 5, 31))
        }

        Dim tran = Me.mSQLite.BeginTransaction()
        Try
            Me.mSQLite.SetTransaction(tran).ExecuteQuery(
                "INSERT INTO Zodiac (name, jp_name, from_date, to_date) 
                    VALUES (@Name, @JpName, @FromDate, @ToDate)", Nothing, zodiacs)
            Me.mSQLite.SetTransaction(tran).ExecuteQuery(
                "INSERT INTO Person (name, zodiac, birth_day) 
                    VALUES (@Name, @Zodiac, @BirthDay)", Nothing, persons)

            tran.Commit()
        Catch ex As Exception
            tran.Rollback()
        End Try

        'Dim ansZodiacs As New PrimaryKeyList(Of Zodiac)()
        Dim ansZodiacs = Me.mSQLite.ExecuteCustomRecords(Of Zodiac)(
            "select " &
            "  Person.Name, Person.birth_day, Zodiac.name, Zodiac.jp_name, Zodiac.from_date, Zodiac.to_date " &
            "from Person " &
            "left outer join Zodiac on " &
            "  Person.zodiac = Zodiac.name",
            Function(prm, keys) As Zodiac
                Dim pson = New Person(prm(0).ToString(), prm(2).ToString(), CDate(prm(1)))
                Dim zdicKey = prm(2).ToString()

                Dim registed = keys.SearchValue(zdicKey)
                If registed.hasValue Then
                    registed.value.Persons.Add(pson)
                    Return Nothing
                Else
                    Dim zdic = New Zodiac(zdicKey, prm(3).ToString(), CDate(prm(4)), CDate(prm(5)))
                    zdic.Persons.Add(pson)
                    keys.Regist(zdic, zdicKey)
                    Return zdic
                End If
            End Function
        ).ToDictionary(Function(v) v.Name, Function(v) v)
        Assert.Equal(ansZodiacs("Aries").Persons.Count, 3)
        Assert.Equal(ansZodiacs("Taurus").Persons.Count, 2)
        Assert.Equal(ansZodiacs("Gemini").Persons.Count, 1)

        Dim tran3 = Me.mSQLite.BeginTransaction()
        Try
            Me.mSQLite.SetTransaction(tran3).ExecuteQuery("DROP TABLE Zodiac")
            Me.mSQLite.SetTransaction(tran3).ExecuteQuery("DROP TABLE Person")
            tran3.Commit()
        Catch ex As Exception
            tran3.Rollback()
        End Try
    End Sub

    <Fact>
    Public Async Sub ParameterCheckerTest()
        Dim tran1 = Me.mSQLite.BeginTransaction()
        Try
            Me.mSQLite.SetTransaction(tran1).ExecuteQuery(
"CREATE TABLE Person (
name TEXT,
zodiac TEXT,
birth_day DATETIME NOT NULL,
PRIMARY KEY(name)
)")

            tran1.Commit()
        Catch ex As Exception
            tran1.Rollback()
        End Try

        Dim persons = New Person() {
New Person("佐藤 健", "Aries", New Date(1989, 3, 21)),
New Person("岩城 滉一", "Aries", New Date(1951, 3, 21)),
New Person("大橋 巨泉", "Aries", New Date(1934, 3, 22)),
New Person("阿部 サダヲ", "Taurus", New Date(1970, 4, 23)),
New Person("大沢 樹生", "Taurus", New Date(1969, 4, 20)),
New Person("有吉 弘行", "Gemini", New Date(1974, 5, 31))
}

        Dim tran = Me.mSQLite.BeginTransaction()
        Try
            Me.mSQLite.SetTransaction(tran).ExecuteQuery(
                "INSERT INTO Person (name, zodiac, birth_day) 
                    VALUES (@Name, @Zodiac, @BirthDay)", Nothing, persons)

            tran.Commit()
        Catch ex As Exception
            tran.Rollback()
        End Try

        Dim tbl = Await Me.mSQLite.
            SetParameterChecker(
                Sub(prm)
                    prm.DbType = DbType.String
                End Sub
            ).
            ExecuteTableSync(
                "select * from Person where zodiac = @Zodiac",
                New With {.Zodiac = "Aries"}
            )

        Assert.Equal(tbl.Rows.Count, 3)

        Dim tbl2 = Await Me.mSQLite.
            SetOrderName("Name").
            ExecuteTableSync(
                "select * from Person where name = ?",
                New With {.Name = "阿部 サダヲ"}
            )
        Assert.Equal(tbl2.Rows.Count, 1)

        Dim tran3 = Me.mSQLite.BeginTransaction()
        Try
            Me.mSQLite.SetTransaction(tran3).ExecuteQuery("DROP TABLE Person")
            tran3.Commit()
        Catch ex As Exception
            tran3.Rollback()
        End Try
    End Sub

End Class

Public Class Zodiac
    Public Property Name As String
    Public Property JpName As String
    Public Property FromDate As Date
    Public Property ToDate As Date
    Public Property Persons As New List(Of Person)
    Public Sub New(nm As String, jp As String, frm As Date, tod As Date)
        Me.Name = nm
        Me.JpName = jp
        Me.FromDate = frm
        Me.ToDate = tod
    End Sub
    Public Overrides Function Equals(obj As Object) As Boolean
        With TryCast(obj, Zodiac)
            Return Me.Name = .Name AndAlso
                   Me.JpName = .JpName AndAlso
                   Me.FromDate = .FromDate AndAlso
                   Me.ToDate = .ToDate
        End With
    End Function
End Class

Public Class Person
    Public ReadOnly Property Name As String
    Public ReadOnly Property Zodiac As String
    Public ReadOnly Property BirthDay As Date
    Public Sub New(nm As String, zodc As String, bday As Date)
        Me.Name = nm
        Me.Zodiac = zodc
        Me.BirthDay = bday
    End Sub
End Class