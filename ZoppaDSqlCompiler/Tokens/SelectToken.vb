Option Strict On
Option Explicit On

Imports ZoppaDSqlCompiler.TokenCollection

Namespace Tokens

    ''' <summary>Selectトークン。</summary>
    Public NotInheritable Class SelectToken
        Implements IToken, ICommandToken, IControlToken

        ' 条件式トークン
        Private ReadOnly mToken As List(Of TokenPosition)

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Throw New NotImplementedException("使用できません")
            End Get
        End Property

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        Public ReadOnly Property TokenType As Type Implements IToken.TokenType
            Get
                Return GetType(SelectToken)
            End Get
        End Property

        ''' <summary>条件式トークンリストを取得します。</summary>
        ''' <returns>条件式トークンリスト。</returns>
        Public ReadOnly Property CommandTokens As List(Of TokenPosition) Implements ICommandToken.CommandTokens
            Get
                Return Me.mToken
            End Get
        End Property

        ''' <summary>トークンが空白文字ならば真を返します。</summary>
        ''' <returns>トークンが空白文字ならば真。</returns>
        Public ReadOnly Property IsWhiteSpace As Boolean Implements IToken.IsWhiteSpace
            Get
                Return False
            End Get
        End Property

        ''' <summary>トークンが改行文字ならば真を返します。</summary>
        ''' <returns>トークンが改行文字ならば真。</returns>
        Public ReadOnly Property IsCrLf As Boolean Implements IToken.IsCrLf
            Get
                Return False
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tokens">条件式のトークン。</param>
        Public Sub New(tokens As List(Of TokenPosition))
            Me.mToken = New List(Of TokenPosition)(tokens)
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "Select"
        End Function

    End Class

End Namespace