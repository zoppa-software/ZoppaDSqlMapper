Option Strict On
Option Explicit On

Imports System.Collections.Concurrent
Imports System.Reflection

''' <summary>フォーマット済みログ値クラスです。</summary>
Public Structure ZoppaFormattedLogValues
    Implements IReadOnlyList(Of KeyValuePair(Of String, Object))

    ''' <summary>キャッシュの最大サイズ。</summary>
    Private Const MaxCachedFormatters As Integer = 64

    ''' <summary>NULLフォーマット。</summary>
    Private Const NullFormat As String = "[null]"

    ' フォーマッターのキャッシュ
    Private Shared ReadOnly s_formatters As New SortedDictionary(Of String, ZoppaLogValuesFormatter)()

    ' フォーマッターの履歴
    Private Shared ReadOnly s_history As New Queue(Of String)()

    ' フォーマッター
    Private ReadOnly _formatter As ZoppaLogValuesFormatter

    ' ログ引数リスト
    Private ReadOnly _values As Object()

    ' オリジナルメッセージ
    Public ReadOnly Property OriginalMessage As String

    ''' <summary>ログ出力したクラスを取得します。</summary>
    Public ReadOnly Property LogClass As Type

    ''' <summary>ログ出力したメンバーを取得します。</summary>
    Public ReadOnly Property LogMember As String

    ''' <summary>ログ出力した行位置を取得します。</summary>
    Public ReadOnly Property LineNumber As Integer

    ''' <summary>フォーマッターを取得する。</summary>
    ''' <returns>フォーマッター。</returns>
    Friend ReadOnly Property Formatter As ZoppaLogValuesFormatter
        Get
            Return Me._formatter
        End Get
    End Property

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="format">フォーマット。</param>
    ''' <param name="frame">スタックフレーム。</param>
    ''' <param name="values">値。</param>
    Public Sub New(format As String, callerType As Type, callerMember As String, lineNumber As Integer, ParamArray values() As Object)
        If values?.Length > 0 AndAlso format IsNot Nothing Then
            SyncLock s_formatters
                If s_formatters.ContainsKey(format) Then
                    Me._formatter = s_formatters(format)
                Else
                    If s_formatters.Count >= MaxCachedFormatters Then
                        Dim f = s_history.Dequeue()
                        s_formatters.Remove(f)
                    End If

                    Me._formatter = New ZoppaLogValuesFormatter(format)
                    s_formatters.Add(format, Me._formatter)
                    s_history.Enqueue(format)
                End If
            End SyncLock
        Else
            Me._formatter = Nothing
        End If

        Me.OriginalMessage = If(format, NullFormat)
        Me.LogClass = callerType
        Me.LogMember = callerMember
        Me.LineNumber = lineNumber
        Me._values = values
    End Sub

    ''' <summary>インデクサ。</summary>
    ''' <param name="index">インデックス。</param>
    ''' <returns>値。</returns>
    Default Public ReadOnly Property Item(index As Integer) As KeyValuePair(Of String, Object) Implements IReadOnlyList(Of KeyValuePair(Of String, Object)).Item
        Get
            If index < 0 OrElse index >= Me.Count Then
                Throw New IndexOutOfRangeException(NameOf(index))
            End If

            If index = Me.Count - 1 Then
                Return New KeyValuePair(Of String, Object)("{OriginalFormat}", Me.OriginalMessage)
            End If

            Return Me._formatter.GetValue(Me._values, index)
        End Get
    End Property

    ''' <summary>要素数を取得する。</summary>
    ''' <returns>要素数。</returns>
    Public ReadOnly Property Count As Integer Implements IReadOnlyCollection(Of KeyValuePair(Of String, Object)).Count
        Get
            If Me._formatter Is Nothing Then
                Return 1
            End If
            Return Me._formatter.ValueNames.Count + 1
        End Get
    End Property

    ''' <summary>列挙子を取得する。</summary>
    ''' <returns>列挙子。</returns>
    Public Iterator Function GetEnumerator() As IEnumerator(Of KeyValuePair(Of String, Object)) Implements IEnumerable(Of KeyValuePair(Of String, Object)).GetEnumerator
        For i As Integer = 0 To Me.Count - 1
            Yield Me(i)
        Next
    End Function

    ''' <summary>文字列を取得する。</summary>
    ''' <returns>文字列。</returns>
    Public Overrides Function ToString() As String
        Return Me.LogMessage()
    End Function

    ''' <summary>列挙子を取得する。</summary>
    ''' <returns>列挙子。</returns>
    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    ''' <summary>ログメッセージを作成する。</summary>
    ''' <returns>メッセージ。</returns>
    Public Function LogMessage() As String
        Return If(Me._formatter Is Nothing, Me.OriginalMessage, Me._formatter.Format(Me._values))
    End Function

    ''' <summary>例外を含むログメッセージを作成する。</summary>
    ''' <param name="ex">例外。</param>
    ''' <returns>メッセージ。</returns>
    Public Function LogMessage(ex As Exception) As String
        Dim msg = Me.LogMessage()
        If msg.Trim() <> "" Then
            Return $"{ex}"
        Else
            Return $"{msg}{Environment.NewLine}{ex}"
        End If
    End Function

End Structure
