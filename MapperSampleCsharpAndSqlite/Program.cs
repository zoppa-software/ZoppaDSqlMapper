﻿using Microsoft.Extensions.Logging;
using System.Data.SQLite;
using System.Xml.Linq;
using ZoppaDSql;

using var loggerFactory = ZoppaDSqlManager.CreateZoppaDSqlLogFactory();

using (var sqlite = new SQLiteConnection("Data Source=chinook.db")) {
    sqlite.Open();

    var query =
@"select
  albumid, title, name
from
  albums
inner join artists on
  albums.ArtistId = artists.ArtistId
where
  albums.ArtistId = @seachId
";

    var answer = sqlite.ExecuteRecords<AlbumInfo>(query, new int[]{11, 23}.Select(v => new { seachId = v }));


    var query1 =
@"SELECT
    AlbumId       -- AlbumId
    , Title       -- Title
    , ArtistId    -- ArtistId
FROM
    albums 
WHERE
    ArtistId = @ArtistId 
ORDER BY
    ArtistId";

    var ans = sqlite.ExecuteRecords<Album>(query1, new { ArtistId = 23 });
    foreach (var v in ans) {
        Console.WriteLine("AlbumId={0}, AlbumTitle={1}, ArtistId={2}", v.AlbumId, v.Title, v.ArtistId);
    }

    var query2 =
@"SELECT
    artists.ArtistId,   -- ArtistId
    Name,               -- Name
    AlbumId,            -- AlbumId
    Title               -- Title
FROM
    artists
LEFT OUTER JOIN albums ON
    artists.ArtistId = albums.ArtistId
ORDER BY
    artists.ArtistId";

    var ans2 = sqlite.ExecuteCreateRecords<Artist>(
        query2,
        (fields) => [fields[0]],  // キーを取得
        (registed, fields) => {     // 既に登録されている場合の処理
            if (fields[2] != null) {
                registed.Albums.Add(new Album((long)fields[2], (string)fields[3], (long)fields[0]));
            }
        },
        (fields) => {               // 未登録の場合の処理
            var artist = new Artist((long)fields[0], (string)fields[1]);
            if (fields[2] != null) {
                artist.Albums.Add(new Album((long)fields[2], (string)fields[3], (long)fields[0]));
            }
            return artist;
        }
    );
    foreach (var v in ans2) {
        Console.WriteLine("ArtistId={0}, Name={1}", v.ArtistId, v.Name);
    }

    var query3 =
@"SELECT
    InvoiceId              -- InvoiceId
    , CustomerId           -- CustomerId
    , InvoiceDate          -- InvoiceDate
    , BillingAddress       -- BillingAddress
    , BillingCity          -- BillingCity
    , BillingState         -- BillingState
    , BillingCountry       -- BillingCountry
    , BillingPostalCode    -- BillingPostalCode
    , Total                -- Total
FROM
    invoices
ORDER BY
    CustomerId";
    var ans3 = sqlite.ExecuteTable(query3);

    var query4 =
@"SELECT
    tbl       -- tbl
    , idx     -- idx
    , stat    -- stat
FROM
    sqlite_stat1";
    var ans4 = sqlite.ExecuteObject(query4);

    var query5 =
@"SELECT
    COUNT(*)
FROM
    employees ";
    var ans5 = sqlite.ExecuteDatas<long>(query5);

    var query6 =
@"SELECT
    FirstName,
    LastName
FROM
    customers 
ORDER BY
    FirstName";
    var ans6 = sqlite.ExecuteArrays(query6);

    sqlite.ExecuteQuery("DELETE FROM SampleDB ");

    using (var tran = sqlite.BeginTransaction()) {
        sqlite.SetTransaction(tran);

        using (var cr = new ZoppaLegacyFiles.Csv.CsvStreamReader("sample.csv")) {
            sqlite.ExecuteQuery<CsvData>("INSERT INTO SampleDB (indexno, name) VALUES (@Indexno, @Name)", cr);
        }

        tran.Commit();
    }
}

//record class AlbumInfo(long AlbumId, string Title, string Name);

class AlbumInfo(long albumId, string title, string name)
{
    public long AlbumId { get; } = albumId;

    public string Title { get; } = title;

    public string Name { get; } = name;
}

record class Album(long AlbumId, string Title, long ArtistId);

class Artist(long artistId, string name)
{
    public long ArtistId { get; } = artistId;

    public string Name { get; } = name;

    public List<Album> Albums { get; } = [];
}

record class CsvData(long Indexno, string Name);