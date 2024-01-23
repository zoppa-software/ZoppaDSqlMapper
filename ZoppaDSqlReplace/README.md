# 説明
`Microsoft.Extensions.Logging` に従って抽象化したログ出力機能を提供します。  
これは、[ZoppaDsql](https://www.nuget.org/packages/ZoppaDSql/)を更新した[ZoppaDSqlMapper](https://www.nuget.org/packages/ZoppaDSqlMapper/)における標準ログ機能として使用されています。  
高効率を目指していないため、ログ出力の効率を気にする場合は別のライブラリの使用を検討してください。  
  
``` csharp
using Microsoft.Extensions.Logging;
using System.Runtime.InteropServices;
using ZoppaLoggingExtensions;

// ログ出力のためのファクトリを作成
using var loggerFactory = LoggerFactory.Create(builder => {
    builder.AddZoppaLogging(
        (config) => config.MinimumLogLevel = LogLevel.Trace
    );
    builder.SetMinimumLevel(LogLevel.Trace);
});

ILogger logger = loggerFactory.CreateLogger<Program>();

// ログ出力の例
using (logger.BeginScope("scope start {a}", 100)) {
    logger.ZLog<Program>().LogDebug(1, "Does this line get hit? {h} {b}", 100, 200);
    using (logger.ZLog<Program>().BeginScope("2")) {
        logger.ZLog<Program>().LogDebug(3, "Nothing to see here.");
        logger.ZLog<Program>().LogDebug(5, "Warning... that was odd.");
    }
    logger.ZLog<Program>().LogDebug(7, "Oops, there was an error.");
}
logger.ZLog<Program>().LogDebug(5, "== 120.");
```

`Microsoft.Extensions.Logging` の初期化ルールに従い、`ILoggingBuilder` に `AddZoppaLogging` の拡張メソッドを使ってログの出力先の設定を行ってください。  
`AddZoppaLogging` には、`ZoppaLoggingConfiguration` を設定するための `Action<ZoppaLoggingConfiguration>` を渡すことができます。  
`ZoppaLoggingConfiguration` には、出力ファイルを指定する `DefaultLogFile` と、カテゴリごとの出力ファイルを指定する `LogFileByCategory` があります。
**注意** ログレベルを指定する `MinimumLogLevel` は `ILoggingBuilder` の `SetMinimumLevel`の値以上を指定するようにしてください。

汎用ホストを使用する例を以下に示します。  
``` vb
    Private Shared Async Function MainAsync(args As String()) As Task
        ' 汎用ホストを作成
        Dim builder = Microsoft.Extensions.Hosting.Host.CreateApplicationBuilder(args)

        ' 汎用ホストを使用して初期化
        builder.Logging.ClearProviders()
        builder.Logging.AddZoppaLogging(
            Sub(config)
                config.MinimumLogLevel = LogLevel.Trace
            End Sub
        )
        builder.Logging.SetMinimumLevel(LogLevel.Trace)

        ' 汎用ホストをビルド、ログ出力
        Using host = builder.Build()
            Dim loggerFactory = host.Services.GetRequiredService(Of ILoggerFactory)()

            Dim logger = loggerFactory.CreateLogger("Test")

            logger.ZLog(Of Program).LogDebug(1, "Does this line get hit?")
            logger.ZLog(Of Program).LogInformation(3, "Nothing to see here.")
            logger.ZLog(Of Program).LogWarning(5, "Warning... that was odd.")
            logger.ZLog(Of Program).LogError(7, "Oops, there was an error.")
            logger.ZLog(Of Program).LogTrace(5, "== 120.")

            Await host.RunAsync()
        End Using
    End Function
```
汎用ホストを使用する場合は、`appsettings.json`で出力するファイルを指定することができます。
``` json
{
  "Logging": {
	"ZoppaLogging": {
	  "DefaultLogFile": "log.txt",
	  "LogFileByCategory": {
		"Test": "test.txt"
	  }
	}
  }
}
```

最後に `ILogger` の拡張メソッド `ZLog` を使用すると呼び出し元のメソッド名と行番号を出力することができます。
