Option Strict On
Option Explicit On

Imports ZoppaDSqlCompiler.TokenCollection

Namespace Tokens

    ''' <summary>トークンインターフェイス。</summary>
    Public Interface IToken

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        ReadOnly Property Contents() As Object

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        ReadOnly Property TokenType As Type

        ''' <summary>トークンが空白文字ならば真を返します。</summary>
        ''' <returns>トークンが空白文字ならば真。</returns>
        ReadOnly Property IsWhiteSpace As Boolean

        ''' <summary>トークンが改行文字ならば真を返します。</summary>
        ''' <returns>トークンが改行文字ならば真。</returns>
        ReadOnly Property IsCrLf As Boolean

    End Interface

    ''' <summary>命令付トークンインターフェイス。</summary>
    Friend Interface ICommandToken

        ''' <summary>命令トークンリストを取得します。</summary>
        ''' <returns>命令トークンリスト。</returns>
        ReadOnly Property CommandTokens As List(Of TokenPosition)

    End Interface

    ''' <summary>コントロールマーカーインターフェイス。</summary>
    Friend Interface IControlToken

    End Interface

End Namespace

