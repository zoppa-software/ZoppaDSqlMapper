Option Strict On
Option Explicit On

Namespace Tokens

    ''' <summary>Trimトークン。</summary>
    Public NotInheritable Class TrimToken
        Implements IToken, IControlToken

        ''' <summary>末尾からトリムする文字列を返します。</summary>
        Public ReadOnly Property TrimString As String

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Throw New NotImplementedException("使用できません")
            End Get
        End Property

        ''' <summary>トークン名を取得する。</summary>
        ''' <returns>トークン名。</returns>
        Public ReadOnly Property TokenName As String Implements IToken.TokenName
            Get
                Return NameOf(TrimToken)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        Public Sub New()
            Me.TrimString = ""
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="trimStr">末尾からトリム文字。</param>
        Public Sub New(trimStr As String)
            Me.TrimString = trimStr
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "Trim"
        End Function

    End Class

End Namespace
