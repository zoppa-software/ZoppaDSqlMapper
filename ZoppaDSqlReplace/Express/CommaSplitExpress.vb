Option Strict On
Option Explicit On

Imports ZoppaDSqlReplace.Environments
Imports ZoppaDSqlReplace.Tokens

Namespace Express

    ''' <summary>カンマ分割式。</summary>
    Public NotInheritable Class CommaSplitExpress
        Implements IExpression

        ' 配列式
        Private ReadOnly mTms As List(Of IExpression)

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="list">式リスト。</param>
        Public Sub New(list As List(Of IExpression))
            Me.mTms = New List(Of IExpression)(list)
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Return New ArrayToken(Me.mTms.Select(Function(x) x.Executes(env)).ToArray())
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:and"
        End Function

    End Class

End Namespace
