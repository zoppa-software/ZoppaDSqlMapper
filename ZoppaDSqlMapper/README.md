# 説明
これは、[ZoppaDSql](https://www.nuget.org/packages/ZoppaDSql/)を更新した動的にSQLを置き換えるライブラリです。  
  
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
以上、簡単な説明となります。**ライブラリの詳細は**[Githubのページ](https://github.com/zoppa-software/ZoppaDSqlMapper)**を参照してください。**

# 更新について
* 1.0.0 初回リリース