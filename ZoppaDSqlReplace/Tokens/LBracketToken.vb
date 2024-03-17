Option Strict On
Option Explicit On

Namespace Tokens

    ''' <summary>左ブラケットトークン。</summary>
    Public NotInheritable Class LBracketToken
        Implements IToken

        ''' <summary>遅延インスタンス生成プロパティ。</summary>
        Private Shared ReadOnly Property LazyInstance() As New Lazy(Of LBracketToken)(Function() New LBracketToken())

        ''' <summary>唯一のインスタンスを返します。</summary>
        Public Shared ReadOnly Property Value() As LBracketToken
            Get
                Return LazyInstance.Value
            End Get
        End Property

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
                Return GetType(LBracketToken)
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
        Private Sub New()

        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "["
        End Function

    End Class

End Namespace

