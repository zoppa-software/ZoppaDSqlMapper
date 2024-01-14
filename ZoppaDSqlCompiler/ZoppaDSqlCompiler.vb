Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports Microsoft.Extensions.Logging
Imports Microsoft.Extensions.DependencyInjection
Imports ZoppaLoggingExtensions
Imports System.Text
Imports ZoppaDSqlCompiler.Tokens
Imports ZoppaDSqlCompiler.Environments
Imports ZoppaDSqlCompiler.TokenCollection
Imports System.Linq.Expressions

''' <summary>DSql APIモジュール。</summary>
Public Module ZoppaDSqlCompiler

    ''' <summary>解析したSQLのキャッシュサイズ。</summary>
    Public Const MAX_HISTORY_SIZE As Integer = 5

    ' サービスプロバイダー
    Private _provider As IServiceProvider = Nothing

    ' ログファクトリ
    Private _logFactory As ILoggerFactory = Nothing

    ' ロガー
    Private ReadOnly _logger As New Lazy(Of ILogger)(
        Function()
            If _provider IsNot Nothing Then
                Return _provider.GetService(Of ILoggerFactory)()?.CreateLogger("ZoppaDSqlCompiler")
            ElseIf _logFactory IsNot Nothing Then
                Return _logFactory.CreateLogger("ZoppaDSqlCompiler")
            Else
                Return Nothing
            End If
        End Function
    )

    ' コマンド履歴
    Private _cmdHistory As New List(Of (String, List(Of TokenPosition)))

    ''' <summary>ログ出力オブジェクトを取得します。</summary>
    ''' <returns>ログ出力オブジェクト。</returns>
    Friend ReadOnly Property Logger As Lazy(Of ILogger)
        Get
            Return _logger
        End Get
    End Property

#Region "logging"

    ''' <summary>ログ出力を開始します。</summary>
    ''' <param name="defaultLogFile">デフォルトログファイル名。</param>
    ''' <param name="isConsole">コンソールにログを出力するかどうか。</param>
    ''' <param name="encodeName">出力エンコード名。</param>
    ''' <param name="maxLogSize">最大ログサイズ。</param>
    ''' <param name="logGeneration">最大ログ世代数。</param>
    ''' <param name="minimumLogLevel">最小ログ出力レベル。</param>
    ''' <param name="switchByDay">日付が変わったら切り替えるフラグ。</param>
    ''' <param name="cacheLimit">キャッシュに保存するログ行数のリミット値。</param>
    ''' <returns>ログ出力オブジェクト。</returns>
    Public Function CreateZoppaDSqlLogFactory(Optional defaultLogFile As String = "zoppa_dsql.log",
                                              Optional isConsole As Boolean = False,
                                              Optional encodeName As String = "utf-8",
                                              Optional maxLogSize As Integer = 30 * 1024 * 1024,
                                              Optional logGeneration As Integer = 10,
                                              Optional minimumLogLevel As LogLevel = LogLevel.Debug,
                                              Optional switchByDay As Boolean = True,
                                              Optional cacheLimit As Integer = 1000) As ILoggerFactory
        Dim loggerFactory = Microsoft.Extensions.Logging.LoggerFactory.Create(
            Sub(builder)
                builder.AddZoppaLogging(
                    Sub(config)
                        config.DefaultLogFile = defaultLogFile
                        config.IsConsole = isConsole
                        config.EncodeName = encodeName
                        config.MaxLogSize = maxLogSize
                        config.LogGeneration = logGeneration
                        config.MinimumLogLevel = minimumLogLevel
                        config.SwitchByDay = switchByDay
                        config.CacheLimit = cacheLimit
                    End Sub
                )
                builder.SetMinimumLevel(minimumLogLevel)
            End Sub
        )

        _logFactory = loggerFactory

        Using scope = _logger.Value?.BeginScope("log setting")
            _logger.Value?.LogInformation($"output log file : {defaultLogFile}")
            _logger.Value?.LogInformation($"is console out : {isConsole}")
            _logger.Value?.LogInformation($"encode name : {encodeName}")
            _logger.Value?.LogInformation($"max log file size : {maxLogSize}")
            _logger.Value?.LogInformation($"log generation : {logGeneration}")
            _logger.Value?.LogInformation($"minimum log level : {minimumLogLevel}")
            _logger.Value?.LogInformation($"switch by day : {switchByDay}")
            _logger.Value?.LogInformation($"output cache limit : {cacheLimit}")
        End Using

        Return loggerFactory
    End Function

    ''' <summary>外部サービスプロバイダーからログ出力を行います。</summary>
    ''' <param name="provider">サービスプロバイダー。</param>
    ''' <returns>サービスプロバイダー。</returns>
    <Extension>
    Public Function SetZoppaDSqlLogProvider(provider As IServiceProvider) As IServiceProvider
        If _provider Is Nothing Then
            _provider = provider
            Using scope = _logger.Value?.BeginScope("log setting")
                _logger.Value?.LogInformation("use other log by service provider")
            End Using
        End If
        Return _provider
    End Function

    ''' <summary>外部ログファクトリからログ出力を行います。</summary>
    ''' <param name="factory">ログファクトリ。</param>
    ''' <returns>ログファクトリ。</returns>
    <Extension>
    Public Function SetZoppaDSqlLogFactory(factory As ILoggerFactory) As ILoggerFactory
        If _logFactory Is Nothing Then
            _logFactory = factory
            Using scope = _logger.Value?.BeginScope("log setting")
                _logger.Value?.LogInformation("use other log by logger factory")
            End Using
        End If
        Return _logFactory
    End Function

    ''' <summary>パラメータをログ出力します。</summary>
    ''' <param name="parameter">パラメータ。</param>
    Private Sub LoggingParameter(parameter As Object)
        If parameter IsNot Nothing Then
            _logger.Value?.LogDebug("Parameters")
            Dim props = parameter.GetType().GetProperties()
            For Each prop In props
                Dim v = prop.GetValue(parameter)
                _logger.Value?.LogDebug("・{prop.Name}={GetValueString(v)} ({prop.PropertyType})", prop.Name, GetValueString(v), prop.PropertyType)
            Next
        End If
    End Sub

    ''' <summary>オブジェクトの文字列表現を取得します。</summary>
    ''' <param name="value">オブジェクト。</param>
    ''' <returns>文字列表現。</returns>
    Private Function GetValueString(value As Object) As String
        If TypeOf value Is IEnumerable Then
            Dim buf As New StringBuilder()
            For Each v In CType(value, IEnumerable)
                If buf.Length > 0 Then buf.Append(", ")
                buf.Append(v)
            Next
            Return buf.ToString()
        Else
            Return If(value?.ToString(), "[null]")
        End If
    End Function

    Private Sub LoggingTokens(tokens As List(Of TokenPosition))
        Logger.Value?.LogTrace("token count : {tokens}", tokens.Count)
        If Logger.Value?.IsEnabled(LogLevel.Trace) Then
            For i As Integer = 0 To tokens.Count - 1
                Logger.Value?.LogTrace("・{tokens(i).Token} ({tokens(i).Position})", tokens(i), tokens(i).Position)
            Next
        End If
    End Sub

#End Region

    ''' <summary>動的SQLをコンパイルします。</summary>
    ''' <param name="sqlQuery">動的SQL。</param>
    ''' <param name="parameter">動的SQL、クエリパラメータ用の情報。</param>
    ''' <returns>コンパイル結果。</returns>
    <Extension>
    Public Function Compile(sqlQuery As String, Optional parameter As Object = Nothing) As String
        Using scope = _logger.Value?.BeginScope(NameOf(Compile))
            Try
                Logger.Value?.LogDebug("compile sql : {sqlQuery}", sqlQuery)
                LoggingParameter(parameter)

                ' トークンリストを取得
                Dim tokens = GetNewOrHistory(sqlQuery, AddressOf LexicalAnalysis.SplitQueryToken)
                LoggingTokens(tokens)

                ' 評価
                Dim ans = ParserAnalysis.Replase(tokens, New EnvironmentObjectValue(parameter))
                Logger.Value?.LogDebug("answer compile sql : {ans}", ans)

            Catch ex As Exception
                _logger.Value?.LogError("message:{ex.Message} stack trace:{ex.StackTrace}", ex.Message, ex.StackTrace)
                Throw
            End Try
            'Try
            '    LoggingDebug($"Compile SQL : {sqlQuery}")
            '    LoggingParameter(parameter)
            '    Dim ans = ParserAnalysis.Replase(sqlQuery, parameter)
            '    LoggingDebug($"Answer SQL : {ans}")
            '    Return ans
            'Catch ex As Exception
            '    LoggingError(ex.Message)
            '    LoggingError(ex.StackTrace)
            '    Throw
            'End Try
        End Using
    End Function

    ''' <summary>あれば履歴からトークンリストを取得、なければ新規作成して履歴に追加します。</summary>
    ''' <param name="sqlQuery">評価する文字列。</param>
    ''' <returns>トークンリスト。</returns>
    Private Function GetNewOrHistoryByCompile(sqlQuery As String) As List(Of TokenPosition)
        SyncLock _cmdHistory
            ' 履歴にあればぞれを返す
            For i As Integer = 0 To _cmdHistory.Count - 1
                If _cmdHistory(i).Item1 = sqlQuery Then
                    Return _cmdHistory(i).Item2
                End If
            Next

            ' 履歴になければ新規作成
            Dim tokens = LexicalAnalysis.SplitToken(sqlQuery)
            _cmdHistory.Add((sqlQuery, tokens))

            If _cmdHistory.Count > MAX_HISTORY_SIZE Then
                _cmdHistory.RemoveAt(0)
            End If

            Return tokens
        End SyncLock
    End Function

    ''' <summary>引数の文字列を評価して値を取得します。</summary>
    ''' <param name="expression">評価する文字列。</param>
    ''' <param name="parameter">環境値。</param>
    ''' <returns>評価結果。</returns>
    <Extension>
    Public Function Executes(expression As String, Optional parameter As Object = Nothing) As IToken
        Using scope = _logger.Value?.BeginScope(NameOf(Executes))
            Try
                Logger.Value?.LogDebug("executes expression : {expression}", expression)
                LoggingParameter(parameter)

                ' トークンリストを取得
                Dim tokens = GetNewOrHistory(expression, AddressOf LexicalAnalysis.SplitToken)
                LoggingTokens(tokens)

                ' 評価
                Dim ans = ParserAnalysis.Executes(tokens, New EnvironmentObjectValue(parameter))
                Logger.Value?.LogDebug("answer expression : {ans}", ans)

                Return ans

            Catch ex As Exception
                _logger.Value?.LogError("message:{ex.Message} stack trace:{ex.StackTrace}", ex.Message, ex.StackTrace)
                Throw
            End Try
        End Using
    End Function

    ''' <summary>あれば履歴からトークンリストを取得、なければ新規作成して履歴に追加します。</summary>
    ''' <param name="sqlQuery">評価する文字列。</param>
    ''' <returns>トークンリスト。</returns>
    Private Function GetNewOrHistory(sqlQuery As String, splitFunc As Func(Of String, List(Of TokenPosition))) As List(Of TokenPosition)
        SyncLock _cmdHistory
            ' 履歴にあればぞれを返す
            For i As Integer = 0 To _cmdHistory.Count - 1
                If _cmdHistory(i).Item1 = sqlQuery Then
                    Return _cmdHistory(i).Item2
                End If
            Next

            ' 履歴になければ新規作成
            Dim tokens = splitFunc(sqlQuery)
            _cmdHistory.Add((sqlQuery, tokens))

            If _cmdHistory.Count > MAX_HISTORY_SIZE Then
                _cmdHistory.RemoveAt(0)
            End If

            Return tokens
        End SyncLock
    End Function

End Module
