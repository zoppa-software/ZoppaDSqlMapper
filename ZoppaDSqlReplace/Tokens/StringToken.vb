Option Strict On
Option Explicit On

Namespace Tokens

    ''' <summary>文字列トークン。</summary>
    Public NotInheritable Class StringToken
        Implements IToken

        ' 文字列
        Private ReadOnly mValue As String

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Return Me.mValue
            End Get
        End Property

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        Public ReadOnly Property TokenType As Type Implements IToken.TokenType
            Get
                Return GetType(StringToken)
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
        ''' <param name="value">文字列。</param>
        Public Sub New(value As String)
            Me.mValue = value
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return Me.mValue
        End Function

    End Class

End Namespace
