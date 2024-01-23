Option Strict On
Option Explicit On

Imports System.Threading
Imports Microsoft.Extensions.Logging
Imports ZoppaLoggingExtensions.ZoppaLoggingLogger

''' <summary>ログ出力クラスです。</summary>
Public NotInheritable Class ZoppaLoggingLogger
    Implements ILogger

    ' ミューティックス
    Private ReadOnly _mutex As Mutex

    ' 内部ロガー
    Private _logger As LocalLogger

    ''' <summary>カテゴリ名です。</summary>
    Public ReadOnly CategoryName As String

    ''' <summary>ログ設定です。</summary>
    Public ReadOnly Config As Func(Of ZoppaLoggingConfiguration)

    ''' <summary>スコープリストです。</summary>
    Public ReadOnly Scopes As New List(Of ScopeBase)

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="categoryName">カテゴリ名。</param>
    ''' <param name="config">ログ設定。</param>
    Public Sub New(categoryName As String, config As Func(Of ZoppaLoggingConfiguration))
        Me.CategoryName = categoryName
        Me.Config = config

        Me._mutex = New Mutex(False, $"Global\zoppalog_{categoryName}")
    End Sub

    ''' <summary>スコープを開始します。</summary>
    ''' <typeparam name="TState">ステート型。</typeparam>
    ''' <param name="state">ステート。</param>
    ''' <returns>スコープ。</returns>
    Public Function BeginScope(Of TState)(state As TState) As IDisposable Implements ILogger.BeginScope
        Return New Scope(Of TState)(Me.Scopes, state)
    End Function

    ''' <summary>指定されたログレベルが有効かどうかを返します。</summary>
    ''' <param name="logLevel">ログレベル。</param>
    ''' <returns>有効ならば真。</returns>
    Public Function IsEnabled(logLevel As LogLevel) As Boolean Implements ILogger.IsEnabled
        Return (logLevel >= Me.Config().MinimumLogLevel)
    End Function

    ''' <summary>ログ出力します。</summary>
    ''' <typeparam name="TState">ステート型。</typeparam>
    ''' <param name="logLevel">ログレベル。</param>
    ''' <param name="eventId">イベントID。</param>
    ''' <param name="state">ステート。</param>
    ''' <param name="exception">例外。</param>
    ''' <param name="formatter">フォーマッター。</param>
    Public Sub Log(Of TState)(logLevel As LogLevel,
                              eventId As EventId,
                              state As TState,
                              exception As Exception,
                              formatter As Func(Of TState, Exception, String)) Implements ILogger.Log
        If Me._logger Is Nothing Then
            Me._mutex.WaitOne()
            Try
                Me._logger = LocalLogger.Create(Me.CategoryName, Me.Config)
            Finally
                Me._mutex.ReleaseMutex()
            End Try
        End If

        If Me.IsEnabled(logLevel) Then
            Me._logger.Stock(New LogData(Of TState)(Me.CategoryName, logLevel, eventId, Me.Scopes, state, exception, formatter))
            Me._logger.Write(Me._mutex)
        End If
    End Sub

    ''' <summary>ログ出力終了を待機します。</summary>
    Public Sub WaitFinish()
        Me._logger?.WaitFinish()
    End Sub

    ''' <summary>スコープ基底クラスです。</summary>
    Public MustInherit Class ScopeBase
        Implements IDisposable

        ' スコープリスト
        Private ReadOnly _scopes As List(Of ScopeBase)

        ''' <summary>スコープ名を取得します。</summary>
        Public MustOverride ReadOnly Property ScopeName As String

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="scopes">スコープリスト。</param>
        Protected Sub New(scopes As List(Of ScopeBase))
            Me._scopes = scopes
            Me._scopes.Add(Me)
        End Sub

        ''' <summary>スコープを終了します。</summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Me._scopes.Remove(Me)
        End Sub

    End Class

    ''' <summary>スコープクラスです。</summary>
    ''' <typeparam name="TState">ステート型。</typeparam>
    Private NotInheritable Class Scope(Of TState)
        Inherits ScopeBase

        ' ステート
        Private ReadOnly _state As TState

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="scopes">スコープリスト。</param>
        ''' <param name="state">ステート。</param>
        Public Sub New(scopes As List(Of ScopeBase), state As TState)
            MyBase.New(scopes)
            Me._state = state
        End Sub

        ''' <summary>スコープ名を取得します。</summary>
        Public Overrides ReadOnly Property ScopeName As String
            Get
                Return Me._state.ToString()
            End Get
        End Property

    End Class

    ''' <summary>書き込む情報。</summary>
    Public NotInheritable Class LogData(Of TState)
        Implements LocalLogger.ILogData

        ''' <summary>出力時刻を取得する。</summary>
        Public ReadOnly Property OutTime As Date

        ''' <summary>カテゴリ名を取得する。</summary>
        Public ReadOnly Property CategoryName As String

        ''' <summary>ログレベルを取得する。</summary>
        Public ReadOnly Property LogLv As LogLevel Implements LocalLogger.ILogData.LogLv

        ''' <summary>イベントIDを取得する。</summary>
        Public ReadOnly Property EvId As EventId

        ''' <summary>スコープ文字列を取得する。</summary>
        Public ReadOnly Property ScopeString As String
            Get
                Dim res As String = ""
                For Each s In Me.Scopes
                    res &= $"|{s.ScopeName}"
                Next
                Return res
            End Get
        End Property

        ''' <summary>スコープリストを取得する。</summary>
        Public ReadOnly Property Scopes As List(Of ZoppaLoggingLogger.ScopeBase)

        ''' <summary>状態を取得する。</summary>
        Public ReadOnly Property State As TState

        ''' <summary>例外を取得する。</summary>
        Public ReadOnly Property Excep As Exception

        ''' <summary>フォーマッタを取得する。</summary>
        Public ReadOnly Property Formatter As Func(Of TState, Exception, String)

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="cateNm">カテゴリ名。</param>
        ''' <param name="lv">ログレベル。</param>
        ''' <param name="eid">イベントID。</param>
        ''' <param name="msg">メッセージ。</param>
        ''' <param name="scopes">スコープリスト。</param>
        ''' <param name="state">状態。</param>
        ''' <param name="exception">例外。</param>
        ''' <param name="formatter">フォーマッタ。</param>
        Public Sub New(cateNm As String,
                       lv As LogLevel,
                       eid As EventId,
                       scopes As List(Of ZoppaLoggingLogger.ScopeBase),
                       state As TState,
                       exception As Exception,
                       formatter As Func(Of TState, Exception, String))
            Me.OutTime = Date.Now
            Me.CategoryName = cateNm
            Me.LogLv = lv
            Me.EvId = eid
            Me.Scopes = New List(Of ZoppaLoggingLogger.ScopeBase)(scopes)
            Me.State = state
            Me.Excep = exception
            Me.Formatter = formatter
        End Sub

        Public Function GetFormatMessage() As String Implements LocalLogger.ILogData.GetFormatMessage
            Try
                Dim msg = Me.Formatter(Me.State, Me.Excep)
                If TypeOf Me.State Is ZoppaFormattedLogValues Then
                    Dim stat As Object = Me.State
                    With CType(stat, ZoppaFormattedLogValues)
                        Return $"[{ Me.OutTime:yyyy/M/d H:mm:ss}|{ Me.LogLv}|{ Me.CategoryName}|{ Me.EvId}{ Me.ScopeString}|{ .LogClass.Name}:{ .LogMember}:{ .LineNumber}] { Me.Formatter(Me.State, Me.Excep)}"
                    End With
                Else
                    Return $"[{ Me.OutTime:yyyy/M/d H:mm:ss}|{ Me.LogLv}|{ Me.CategoryName}|{ Me.EvId}{ Me.ScopeString}] { Me.Formatter(Me.State, Me.Excep)}"
                End If
            Catch ex As FormatException
                If TypeOf Me.State Is ZoppaFormattedLogValues Then
                    Dim stat = CType(CObj(Me.State), ZoppaFormattedLogValues)
                    Return $"[{ Me.OutTime:yyyy/M/d H:mm:ss}|{ Me.LogLv}|{ Me.CategoryName}|{ Me.EvId}{ Me.ScopeString}] format error! {stat.OriginalMessage}:{ex}"
                Else
                    Throw
                End If
            End Try
        End Function

    End Class

End Class
