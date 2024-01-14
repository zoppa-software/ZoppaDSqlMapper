﻿Option Strict On
Option Explicit On

Imports ZoppaDSqlCompiler.TokenCollection

Namespace Tokens

    ''' <summary>ForEachトークン。</summary>
    Public NotInheritable Class ForEachToken
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

        ''' <summary>トークン名を取得する。</summary>
        ''' <returns>トークン名。</returns>
        Public ReadOnly Property TokenName As String Implements IToken.TokenName
            Get
                Return NameOf(ForEachToken)
            End Get
        End Property

        ''' <summary>条件式トークンリストを取得します。</summary>
        ''' <returns>条件式トークンリスト。</returns>
        Public ReadOnly Property CommandTokens As List(Of TokenPosition) Implements ICommandToken.CommandTokens
            Get
                Return Me.mToken
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tokens">ループ変数のトークン。</param>
        Public Sub New(tokens As List(Of TokenPosition))
            Me.mToken = tokens
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "For Each"
        End Function

    End Class

End Namespace

