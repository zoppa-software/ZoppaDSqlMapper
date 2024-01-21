using Microsoft.Extensions.Logging;
using System.Data.SQLite;
using System.Xml.Linq;
using ZoppaDSql;

using var loggerFactory = ZoppaDSqlManager.CreateZoppaDSqlLogFactory();

using (var sqlite = new SQLiteConnection("Data Source=chinook.db")) {
    sqlite.Open();

    var query =
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

    var ans = sqlite.ExecuteRecords<Album>(query, new { ArtistId = 23 });
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

    var ans2 = sqlite.ExecuteCustomRecords<Artist>(
        query2, 
        (valus, primaries) => {
            var registed = primaries.SearchValue((long?)valus[0]);
            if (registed.hasValue) {
                if (valus[2] != null) {
                    registed.value.Albums.Add(new Album((long)valus[2], (string)valus[3], (long)valus[0]));
                }
                return null!;
            }
            else {
                var artist = new Artist((long)valus[0], (string)valus[1]);
                if (valus[2] != null) {
                    artist.Albums.Add(new Album((long)valus[2], (string)valus[3], (long)valus[0]));
                }
                primaries.Regist(artist, artist.ArtistId);
                return artist;
            }
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

record class Album(long AlbumId, string Title, long ArtistId);

class Artist(long artistId, string name)
{
    public long ArtistId { get; } = artistId;

    public string Name { get; } = name;

    public List<Album> Albums { get; } = [];
}

record class CsvData(long Indexno, string Name);