Option Strict On
Option Explicit On

''' <summary>データベース操作例外。</summary>
Public Class ZoppaDSqlException
    Inherits Exception

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="message">例外メッセージ。</param>
    Public Sub New(message As String)
        MyBase.New(message)
    End Sub

End Class
