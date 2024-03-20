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
  