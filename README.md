# ZoppaDSqlMapper
SQL文の内部に制御文を埋め込む方式で動的SQLを行うライブラリです。

以下のようにSQLを動的に置き換えます。  
  
``` vb
Dim answer = "" &
"select * from employees 
{trim}
where
    {trim both}
        {if empNo}emp_no < 20000{/if} and
        {trim}
            ({trim both}{if first_name}first_name like 'A%'{/if} or {if gender}gender = 'F'{/if}{/trim})
        {/trim}
    {/trim}
{/trim}
limit 10".Replase(New With {.empNo = (i And 1) = 0, .first_name = (i And 2) = 0, .gender = (i And 4) = 0})
```
  
SQLを記述した文字列内に `{}` で囲まれた文が制御文になります。  
**メモ** `ZoppaDSql`では `trim` で `and` や `or` を構文解析して、おかしな構文を自動的に削除していましたが、`ZoppaDSqlMapper` では `trim` は囲んだ範囲の後文字か前後の文字列(`trim both`の場合)を削除するだけになります。
拡張メソッド `Replase` にパラメータとなるクラスを引き渡して実行するとクラスのプロパティを読み込み、SQLを動的に構築します。  
パラメータは全て `False` なので、実行結果は以下のようになります。  
``` sql
select * from employees 
limit 10
```
  
以下の例は、パラメータの`seachId`が 0以外ならば、`ArtistId`が`seachId`と等しいことという動的SQLを `query`変数に格納しました。
``` csharp 
var query =
@"select
  albumid, title, name
from
  albums
inner join artists on
  albums.ArtistId = artists.ArtistId
{trim}
where
  {if seachId <> 0} albums.ArtistId = @seachId{/if}
{/trim}";
```
次に、`SQLite`の`IDbConnection`の実装である`SQLiteConnection`を`Open`した後、`ExecuteRecordsSync`拡張メソッドを実行すると`SQL`の実行結果が`AlbumInfo`クラスのリストで取得できます。  
``` csharp 
using (var sqlite = new SQLiteConnection("Data Source=chinook.db")) {
    sqlite.Open();

    var query =
@"select
  albumid, title, name
from
  albums
inner join artists on
  albums.ArtistId = artists.ArtistId
{trim}
where
  {if seachId <> 0} albums.ArtistId = @seachId{/if}
{/trim}";

    var answer = await sqlite.ExecuteRecordsSync<AlbumInfo>(query, new { seachId = 23 });
    // answerにSQLの実行結果が格納されます
}
```
`AlbumInfo`クラスの実装は以下のとおりです。
マッピングは一般的にはプロパティ、フィールドをマッピングしますが、`ZoppaDSqlMapper`は`SQL`の実行結果の各カラムの型と一致する **コンストラクタ** を検索してインスタンスを生成します。
``` csharp 
// クラスを使用した場合
class AlbumInfo(long albumId, string title, string name)
{
    public long AlbumId { get; } = albumId;

    public string Title { get; } = title;

    public string Name { get; } = name;
}

// recordクラスを使用した場合
record class AlbumInfo(long AlbumId, string Title, string Name);
```

## 特徴
* 素に近いSQL文を動的に変更するため、コードから自動生成されるSQL文より調整が容易です  
* SQLが文字列であるため、プログラム言語による差がなくなります  
* select、insert、update、delete を文字列検索すれはデータベース処理を検索できます  

## 依存関係
ライブラリは .NET Standard 2.0 で記述しています。そのため、.net framework 4.6.1以降、.net core 2.0以降で使用できます。   
以下のライブラリを参照します。
* [ZoppaLoggingExtensions](https://www.nuget.org/packages/ZoppaLoggingExtensions/)
* [ZoppaLegacyFiles](https://www.nuget.org/packages/ZoppaLegacyFiles/)
* [ZoppaDSqlReplace](https://www.nuget.org/packages/ZoppaDSqlReplace/)

## 使い方
### SQL文に置き換え式、制御式を埋め込む
SQL文を部分的に置き換える（置き換え式）、また、部分的に除外するや繰り返すなど（制御式）を埋め込みます。  
埋め込みは `#{` *参照するプロパティ* `}`、`{` *制御式* `}` の形式で記述します。  
  
#### 埋め込み式  
`#{` *参照するプロパティ* `}` を使用すると、`Replase`で引き渡したオブジェクトのプロパティを参照して置き換えます。  
以下は文字列プロパティを参照しています。 `'`で囲まれて出力していることに注目してください。   
``` vb
Dim ans1 = "select * from table1 where column = #{value}".Replase(New With {.value = "値"})
Assert.Equal("select * from table1 where column = '値'", ans1)
```
次に数値プロパティを参照します。  
``` vb
Dim ans2 = "select * from member where age >= #{lowAge} and age <= #{hiAge}".Replase(New With {.lowAge = 12, .hiAge = 50})
Assert.Equal("select * from member where age >= 12 and age <= 50", ans2)
```
次にnullを参照します。  
``` vb
Dim ans3 = "update person set name = #{value}".Replase(New With {.value = Nothing})
Assert.Equal("update person set name = null", ans3)
```
埋め込みたい文字列にはテーブル名など `'` で囲みたくない場面があります。この場合、`!{}` (または `${}`)を使用します。  
``` vb
' テーブル
Dim ans4 = "select * from !{table}".Replase(New With {.table = "sample_table"})
Assert.Equal("select * from sample_table", ans4)

' 条件
Dim ans5 = "select * from table2 where !{condition}".Replase(New With {.condition = "clm1 = '123'"})
Assert.Equal("select * from table2 where clm1 = '123'", ans5)
```
三項演算子を使用した置き換えもサポートしています。  
``` vb
Dim ans6 = "update person set name = #{value <> null ? value : ''}".Replase(New With {.value = Nothing})
Assert.Equal("update person set name = ''", ans6)
```
  
#### 制御式  
SQL文を部分的に除外、または繰り返すなど制御を行います。  
  
* **if文**  
条件が真であるならば、その部分を出力します。  
`{if 条件式}`、`{else if 条件式}`、`{else}`、`{end if}`で囲まれた部分を判定します。  
``` vb
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
```

* **select文**  
一致するケースを出力します。  
``` vb
    Enum Mode
        None = 1
        MyGroup = 2
        Specified = 3
    End Enum

        ' 対象のクエリ
        Dim query = "" &
"SELECT
    *
FROM
    TBL1
{trim}
WHERE
{select mode}
{case 'None'}
{case 'MyGroup'}
    GRP = 0
{else}
    GRP = #{groupNo}
{/select}
{/trim}"
        ' Noneの場合は出力しません（WHEREは{trim}でトリムしました）
        Dim ans1 = query.Replase(New With {.mode = Mode.None})
        Assert.Equal(ans1,
"SELECT
    *
FROM
    TBL1
")
        ' MyGroupの場合は GRP = 0
        Dim ans2 = query.Replase(New With {.mode = Mode.MyGroup})
        Assert.Equal(ans2,
"SELECT
    *
FROM
    TBL1
WHERE
    GRP = 0")
        ' 上記以外は条件式
        Dim ans3 = query.Replase(New With {.mode = Mode.Specified, .groupNo = 100})
        Assert.Equal(ans3,
"SELECT
    *
FROM
    TBL1
WHERE
    GRP = 100")
```
  
  
* **foreach文**  
パラメータの要素分、出力を繰り返します。  
`{foreach 一時変数 in パラメータ(配列など)}`、`{end for}`で囲まれた範囲を繰り返します。一時変数に要素が格納されるので foreachの範囲内で置き換え式を使用して出力してください。    
``` vb
"SELECT
    *
FROM
    customers 
WHERE
    FirstName in ({trim}{foreach nm in names}#{nm}, {end for}{end trim})
".Replase(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
```
上記の例では foreachで囲まれた範囲の文、`#{nm}, `がパラメータ`names`の要素数分("Helena", "Dan", "Aaron")を繰り返して出力します。  
(お気づきのとおり、最後の要素では`,`が不要に出力されます、この`,`を取り除くのが **trim**文です)  
出力は以下のとおりです。  
``` vb
SELECT
    *
FROM
    customers 
WHERE
    FirstName in ('Helena', 'Dan', 'Aaron')
```
ただし、このような場合、通常の置き換え式で配列など繰り返し要素を与えれば `,` で結合して置き換えるように実装しています。  
以下の例は上記と同じ結果を出力します。  
``` vb
SELECT
    *
FROM
    customers 
WHERE
    FirstName in (#{names})
".Replase(New With {.names = New String() {"Helena", "Dan", "Aaron"}})
```
  
* **trim文**  
※ 作成中
  
### SQLクエリを実行し、簡単なマッパー機能を使用してインスタンスを生成する
#### 基本的な使い方
動的に生成したSQL文をDapperやEntity Frameworkで利用することができます。  
また、簡単に利用するために [IDbConnection](https://learn.microsoft.com/ja-jp/dotnet/api/system.data.idbconnection)の拡張メソッド、および、シンプルなマッピング処理を用意しています。  
  
以下の例は、パラメータの`seachId`が`NULL`以外ならば、`ArtistId`が`seachId`と等しいことという動的SQLを `query`変数に格納しました。
``` vb
Dim query = "" &
"select
  albumid, title, name
from
  albums
inner join artists on
  albums.ArtistId = artists.ArtistId
where
  {if seachId <> NULL }albums.ArtistId = @seachId{end if}"
```
次に、SQLiteの`IDbConnection`の実装である`SQLiteConnection`をOpenした後、`ExecuteRecordsSync`拡張メソッドを実行するとSQLの実行結果が`AlbumInfo`クラスのリストで取得できます。  
``` vb
Dim answer As List(Of AlbumInfo)
Using sqlite As New SQLiteConnection("Data Source=chinook.db")
    sqlite.Open()

    answer = Await sqlite.ExecuteRecordsSync(Of AlbumInfo)(query, New With {.seachId = 11})
    ' answerにSQLの実行結果が格納されます
End Using
```
`AlbumInfo`クラスの実装は以下のとおりです。  
マッピングは一般的にはプロパティ、フィールドをマッピングしますが、ZoppaDSqlはSQLの実行結果の各カラムの型と一致する**コンストラクタ**を検索してインスタンスを生成します。  
``` vb
Public Class AlbumInfo
    Public ReadOnly Property AlbumId As Integer
    Public ReadOnly Property AlbumTitle As String
    Public ReadOnly Property ArtistName As String

    Public Sub New(id As Long, title As String, nm As String)
        Me.AlbumId = id
        Me.AlbumTitle = title
        Me.ArtistName = nm
    End Sub
End Class
```
#### SQL実行設定
トランザクション、SQLタイムアウトの設定も拡張メソッドで行います。  
以下はトランザクションの例です。  
``` vb
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
```
トランザクションは`IDbConnection`から適切に取得してください。  
`SetTransaction`という拡張メソッドを用意しているのでトランザクションを与えます、その後はコミット、ロールバックを実行してください。  
拡張メソッドは以下のものがあります。  
| メソッド | 内容 |  
| ---- | ---- | 
| SetTransaction | トランザクションを設定します。 | 
| SetTimeoutSecond | SQLタイムアウトを設定します（秒数）、デフォルト値は30秒です。</br>デフォルト値は `ZoppaDSqlSetting.DefaultTimeoutSecond` で変更してください。 | 
| SetParameterPrepix | SQLパラメータの接頭辞を設定します、デフォルトは `@` です。</br>デフォルト値は `ZoppaDSqlSetting.DefaultParameterPrefix` で変更してください。 | 
| SetCommandType | SQLのコマンドタイプを設定します、デフォルト値は `CommandType.Text` です。 |
| SetParameterChecker | SQLパラメータチェック式を設定します。デフォルト値は `Null` です。</br>デフォルトs式は `ZoppaDSqlSetting.DefaultSqlParameterCheck` で変更してください。 |
| SetOrderName | 指定したプロパティ名の順番でSQLパラメータを作成します。 |
  
位置指定パラメータ`?`では名前がないので`SetOrderName`で設定するプロパティの名前を順に設定します。以下は例です、  
``` vb
Dim answer = Await Me.mSQLite.
    SetOrderName("Name").
    ExecuteTableSync(
        "select * from Person where name = ?",
        New With {.Name = "阿部 サダヲ"}
    )
```
### インスタンス生成をカスタマイズします
検索結果が 多対1 など一つのインスタンスで表現できない場合、インスタンスの生成をカスタマイズする必要があります。  
ZoppaDSqlではインスタンスを生成する式を引数で与えることで対応します。  
以下の例では、`Person`テーブルと`Zodiac`テーブルが多対1の関係で、そのまま`Person`クラスと`Zodiac`クラスに展開します。SQLの実行結果 1レコードで 1つの`Person`クラスを生成し、リレーションキーで`Zodiac`クラスをコレクションに保持して、関連を表現する`Persons`プロパティに追加します。
``` vb
Dim ansZodiacs = Me.mSQLite.ExecuteCreateRecords(Of Zodiac)(
    "select " &
    "  Person.Name, Person.birth_day, Zodiac.name, Zodiac.jp_name, Zodiac.from_date, Zodiac.to_date " &
    "from Person " &
    "left outer join Zodiac on " &
    "  Person.zodiac = Zodiac.name",
    Function(prm) ' Zodiacの主キーを返す
        Return {prm(2)}
    End Function,
    Sub(zdic, prm)
        ' 上記主キーで登録済みのZodiacとSQLの取得結果からインスタンスを生成
        Dim pson = New Person(prm(0).ToString(), prm(2).ToString(), CDate(prm(1)))
        zdic.Persons.Add(pson)
    End Sub,
    Function(prm) As Zodiac
        ' 上記主キーで登録済みのZodiacがないので、SQLの取得結果のみでインスタンスを生成
        Dim pson = New Person(prm(0).ToString(), prm(2).ToString(), CDate(prm(1)))
        Dim zdic = New Zodiac(prm(2).ToString(), prm(3).ToString(), CDate(prm(4)), CDate(prm(5)))
        zdic.Persons.Add(pson)
        Return zdic
    End Function
)
```
  
### パラメータにCSVファイルを与えてSQLクエリを実行します
単体テストなど大量データのインサートにはCSVファイルを読み込むライブラリを使用してインサートします。  
以下の例を参照してください。  
``` csharp
using (var tran = sqlite.BeginTransaction()) {
    sqlite.SetTransaction(tran);

    using (var cr = new ZoppaLegacyFiles.Csv.CsvStreamReader("sample.csv")) {
        sqlite.ExecuteQuery<CsvData>("INSERT INTO SampleDB (indexno, name) VALUES (@Indexno, @Name)", cr);
    }

    tran.Commit();
}

record class CsvData(long Indexno, string Name);
```
  
### 実行結果をDataTable、または DynamicObjectで取得します
SQLを実行した結果を取得するためだけにクラスを定義するとクラスの数が増えて管理するのも大変になることもあります。そのため、`DataTable`型、または動的な型（`DynamicObject`）で取得するメソッドを用意しました。  
  
`DataTable`型で取得する例は以下のようになります。  
Rowsプロパティなど使用して実行結果を取得してください。  
``` vb
Dim tbl = Await Me.mSQLite.ExecuteTableSync(
    "select * from Person where zodiac = @Zodiac",
    New With {.Zodiac = "Aries"}
)
```
  
動的な型（`DynamicObject`）で取得する例は以下のようになります。  
(C#では`dynamic`キーワードで動的な型を宣言します。vb.netでは`Object`です。この例ではわかりやすいようにC#で記述しています)  
``` csharp
using (var sqlite = new SQLiteConnection("Data Source=chinook.db")) {
    sqlite.Open();

    var query = 
@"select
  albumid, title, name
from
  albums
inner join artists on
  albums.ArtistId = artists.ArtistId
{trim}
where
  {if seachId <> NULL }albums.ArtistId = @seachId{end if}
{end trim}";

    var ans = sqlite.ExecuteObject(query, new { seachId = 11 });
    foreach (dynamic v in ans) {
        Console.WriteLine("AlbumId={0}, AlbumTitle={1}, ArtistName={2}", v.albumid, v.title, v.name);
    }
}
```
実行結果を受け取る`foreach`処理で実行結果は`dynamic`で受け取っています。albumid、title、nameは`dynamic`型のメンバーではないですが、`ZoppaDSql`内でselect実行結果の列名から動的に追加したものになり、取得することができます。  
※ 動的な型を使用すると記述は簡潔になりますがメンテナンス性は悪くなると考えられます。使用には注意してください
  
### ログファイル出力機能を有効にします
※ 作成中

### 付属機能
#### 簡単な式を評価し、結果を得ることができます
``` vb
' 数式
Dim ans1 = "(28 - 3) / (2 + 3)".Executes().Contents
Assert.Equal(ans1, 5)

' 比較式
Dim ans2 = "0.1 * 5 <= 0.4".Executes().Contents
Assert.Equal(ans2, False)
```
#### カンマ区切りで文字列を分割できます
`"`のエスケープを考慮して文字列を区切ります。分割した文字列は`CsvItem`構造体に格納されます。`Text`プロパティではエスケープは解除されませんが、`UnEscape`メソッドではエスケープを解除します。  
``` vb
Dim csv = CsvSpliter.CreateSpliter("あ,い,う,え,""お,を""").Split()
Assert.Equal(csv(0).UnEscape(), "あ")
Assert.Equal(csv(1).UnEscape(), "い")
Assert.Equal(csv(2).UnEscape(), "う")
Assert.Equal(csv(3).UnEscape(), "え")
Assert.Equal(csv(4).UnEscape(), "お,を")
Assert.Equal(csv(4).Text, """お,を""")
```

## 注意
Oracle DB を対象にマッパー機能でSQLパラメーターを与えるとき、ODP.NET, Managed Driver の OracleParameter の以下の仕様から日本語の検索が正しくできないと思われます。  
> OracleParameter.DbType に DbType.String を設定した場合、OracleDbType には OracleDbType.Varchar2 が設定されます  
> *OracleDbType.NVarchar2 ではありません*  

上記の仕様が問題になる可能性があるので、`IDbDataParameter` を使用前にチェックする式を設定する機能を追加しました。  
以下の例をご覧ください。  
``` vb
Using ora As New OracleConnection()
    ora.ConnectionString = "接続文字列"
    ora.Open()

    Dim tbl = ora.
        SetParameterPrepix(PrefixType.Colon).
        SetParameterChecker(
            Sub(chk)
                Dim prm = TryCast(chk, Oracle.ManagedDataAccess.Client.OracleParameter)
                If prm?.OracleDbType = OracleDbType.Varchar2 Then
                    prm.OracleDbType = OracleDbType.NVarchar2
                End If
            End Sub).
        ExecuteRecords(Of RFLVGROUP)(
            "select * from GROUP where SYAIN_NO = :SyNo ",
            New With {.SyNo = CType("105055", DbString)}
        )
End Using
```
拡張メソッド `SetParameterChecker` では生成した `IDbDataParameter` を順次引き渡すので、`OracleDbType` を式内で変更しています。
全てのSQLに適用したい場合はデフォルトの式を変更します。  
``` vb
ZoppaDSqlSetting.DefaultSqlParameterCheck =
    Sub(chk)
        Dim prm = TryCast(chk, Oracle.ManagedDataAccess.Client.OracleParameter)
        If prm?.OracleDbType = OracleDbType.Varchar2 Then
            prm.OracleDbType = OracleDbType.NVarchar2
        End If
    End Sub
```
  
## インストール
ソースをビルドして `ZoppaDSql.dll` ファイルを生成して参照してください。  
Nugetにライブラリを公開しています。[ZoppaDSql](https://www.nuget.org/packages/ZoppaDSql/)を参照してください。

## 作成情報
* 造田　崇（zoppa software）
* ミウラ第1システムカンパニー 
* takashi.zouta@kkmiuta.jp

## ライセンス
[apache 2.0](https://www.apache.org/licenses/LICENSE-2.0.html)
