Option Strict On
Option Explicit On
Imports System.Runtime.CompilerServices

Namespace Tokens

    ''' <summary>クエリトークン。</summary>
    Public NotInheritable Class QueryToken
        Implements IToken, IControlToken

        ''' <summary>文字種。</summary>
        Public Enum QueryKind

            ''' <summary>改行。</summary>
            IsCrLf = 0

            ''' <summary>空白文字。</summary>
            IsWhiteSpace = 1

            ''' <summary>その他。</summary>
            IsOther = 2

        End Enum

        ' 出力する文字列
        Private ReadOnly mValue As String

        ' 文字種
        Private ReadOnly mKind As QueryKind

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
                Return GetType(QueryToken)
            End Get
        End Property

        ''' <summary>トークンが空白文字ならば真を返します。</summary>
        ''' <returns>トークンが空白文字ならば真。</returns>
        Public ReadOnly Property IsWhiteSpace As Boolean Implements IToken.IsWhiteSpace
            Get
                Return (Me.mKind = QueryKind.IsWhiteSpace)
            End Get
        End Property

        ''' <summary>トークンが改行文字ならば真を返します。</summary>
        ''' <returns>トークンが改行文字ならば真。</returns>
        Public ReadOnly Property IsCrLf As Boolean Implements IToken.IsCrLf
            Get
                Return (Me.mKind = QueryKind.IsCrLf)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="value">対象文字列。</param>
        ''' <param name="kind">クエリの種類。</param>
        Public Sub New(value As String, kind As QueryKind)
            Me.mValue = value
            Me.mKind = kind
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return Me.mValue
        End Function

        ''' <summary>文字種を取得します。</summary>
        ''' <returns>文字種。</returns>
        Public Shared Function GetCharKind(target As Char) As QueryKind
            If target = vbCr OrElse target = vbLf Then
                Return QueryKind.IsCrLf
            ElseIf Char.IsWhiteSpace(target) Then
                Return QueryKind.IsWhiteSpace
            Else
                Return QueryKind.IsOther
            End If
        End Function

    End Class

End Namespace
