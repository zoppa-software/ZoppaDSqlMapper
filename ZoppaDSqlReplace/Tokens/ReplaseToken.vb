Option Strict On
Option Explicit On

Namespace Tokens

    ''' <summary>置き換えトークン。</summary>
    Public NotInheritable Class ReplaseToken
        Implements IToken

        ' 出力する文字列
        Private ReadOnly mValue As String

        ' SQLエスケープするならば真
        Private ReadOnly mIsEscape As Boolean

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
                Return GetType(ReplaseToken)
            End Get
        End Property

        ''' <summary>SQLエスケープを行うならば真を返す。</summary>
        Public ReadOnly Property IsEscape As Boolean
            Get
                Return Me.mIsEscape
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
        ''' <param name="value">出力する文字列。</param>
        Public Sub New(value As String, isEscape As Boolean)
            Me.mValue = value
            Me.mIsEscape = isEscape
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"Bind {Me.mValue} esc:{Me.mIsEscape}"
        End Function

    End Class

End Namespace