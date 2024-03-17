Option Strict On
Option Explicit On

Imports ZoppaDSqlReplace.Environments
Imports ZoppaDSqlReplace.Tokens

Namespace Express

    ''' <summary>配列値式。</summary>
    Public NotInheritable Class ArrayValueExpress
        Implements IExpression

        ' 対象トークン
        Private ReadOnly mToken As IdentToken

        ' 添字トークン
        Private ReadOnly mIndexToken As ParenExpress

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="token">対象トークン。</param>
        ''' <param name="indexToken">インデックストークン。</param>
        Public Sub New(token As IdentToken, indexToken As ParenExpress)
            If token IsNot Nothing AndAlso indexToken IsNot Nothing Then
                Me.mToken = token
                Me.mIndexToken = indexToken
            Else
                Throw New DSqlAnalysisException("値式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Try
                Dim arr = CType(env.GetValue(If(Me.mToken.Contents?.ToString(), "")), IList)
                Dim idx = Convert.ToInt32(Me.mIndexToken.Executes(env).Contents)
                Return ValueExpress.ConvertValueToken(arr(idx))
            Catch ex As Exception
                Throw New DSqlAnalysisException("配列の値を取得できませんでした", ex)
            End Try
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:array value"
        End Function

    End Class

End Namespace
