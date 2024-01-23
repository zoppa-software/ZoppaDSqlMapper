Option Strict On
Option Explicit On

Imports System.IO
Imports Microsoft.Extensions.Logging

''' <summary>ログ設定クラスです。</summary>
Public NotInheritable Class ZoppaLoggingConfiguration
    Implements ICloneable

    ''' <summary>ログファイル設定リストです。</summary>
    Public Property DefaultLogFile As String = "default.log"

    ''' <summary>ログファイル設定リストです。</summary>
    Public Property LogFileByCategory As Dictionary(Of String, String) = New Dictionary(Of String, String)

    ''' <summary>出力エンコード名です。</summary>
    Public Property EncodeName As String = "utf-8"

    ''' <summary>最大ログサイズです。</summary>
    Public Property MaxLogSize As Integer = 30 * 1024 * 1024

    ''' <summary>最大ログ世代数です。</summary>
    Public Property LogGeneration As Integer = 10

    ''' <summary>最小ログ出力レベルです。</summary>
    Public Property MinimumLogLevel As LogLevel = LogLevel.Debug

    ''' <summary>日付が変わったら切り替えるフラグです。</summary>
    Public Property SwitchByDay As Boolean = True

    ''' <summary>キャッシュに保存するログ行数のリミット値です。</summary>
    Public Property CacheLimit As Integer = 1000

    ''' <summary>コンソールにログを出力します。</summary>
    Public Property IsConsole As Boolean = True

    ''' <summary>設定のクローンを作成します。</summary>
    ''' <returns>クローン。</returns>
    Public Function Clone() As Object Implements ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function

End Class
