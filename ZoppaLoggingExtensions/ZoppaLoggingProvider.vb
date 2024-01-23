Option Strict On
Option Explicit On

Imports System.Collections.Concurrent
Imports Microsoft.Extensions.Logging
Imports Microsoft.Extensions.Options

''' <summary>ログ出力クラスを作成するクラスです。</summary>
<ProviderAlias("ZoppaLogging")>
Public NotInheritable Class ZoppaLoggingProvider
    Implements ILoggerProvider

    ' 現在の設定
    Private _currentConfig As ZoppaLoggingConfiguration

    ' 設定変更時のトークン
    Private _onChangeToken As IDisposable

    ' カテゴリごとのログ出力クラス
    Private ReadOnly _loggers As New ConcurrentDictionary(Of String, ZoppaLoggingLogger)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="config">設定。</param>
    Public Sub New(config As IOptionsMonitor(Of ZoppaLoggingConfiguration))
        Me._currentConfig = config.CurrentValue
        Me._onChangeToken = config.OnChange(Sub(c, o) Me._currentConfig = c)
    End Sub

    ''' <summary>ログ出力クラスを作成します。</summary>
    ''' <param name="categoryName">カテゴリ名。</param>
    ''' <returns>ログ出力クラス。</returns>
    Public Function CreateLogger(categoryName As String) As ILogger Implements ILoggerProvider.CreateLogger
        Return _loggers.GetOrAdd(
            categoryName,
            Function(name) As ZoppaLoggingLogger
                Return New ZoppaLoggingLogger(name, Function() Me.GetCurrentConfig())
            End Function
        )
    End Function

    ''' <summary>現在の設定を取得します。</summary>
    ''' <returns>現在の設定。</returns>
    Private Function GetCurrentConfig() As ZoppaLoggingConfiguration
        Return Me._currentConfig
    End Function

    ''' <summary>リソースの解放を行います。</summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        For Each lg In Me._loggers.Values
            lg?.WaitFinish()
        Next
        Me._loggers.Clear()

        Me._onChangeToken?.Dispose()
        Me._onChangeToken = Nothing
    End Sub

End Class
