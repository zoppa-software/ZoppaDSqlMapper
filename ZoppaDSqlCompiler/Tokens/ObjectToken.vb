Option Strict On
Option Explicit On

Namespace Tokens

    ''' <summary>オブジェクトトークン。</summary>
    Public NotInheritable Class ObjectToken
        Implements IToken

        ' オブジェクト
        Private ReadOnly mValue As Object

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Return Me.mValue
            End Get
        End Property

        ''' <summary>トークン名を取得する。</summary>
        ''' <returns>トークン名。</returns>
        Public ReadOnly Property TokenName As String Implements IToken.TokenName
            Get
                Return NameOf(ObjectToken)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">オブジェクト。</param>
        Public Sub New(value As Object)
            Me.mValue = value
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return If(Me.mValue?.ToString(), "Null")
        End Function

    End Class

End Namespace