Option Strict On
Option Explicit On

''' <summary>解析例外。</summary>
Public Class DSqlAnalysisException
    Inherits Exception

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="message">例外メッセージ。</param>
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="message">例外メッセージ。</param>
    ''' <param name="innerEx">内部例外。</param>
    Public Sub New(message As String, innerEx As Exception)
        MyBase.New(message, innerEx)
    End Sub

End Class
