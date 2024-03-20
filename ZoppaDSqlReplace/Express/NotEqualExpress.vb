Option Strict On
Option Explicit On

Imports ZoppaDSqlReplace.Environments
Imports ZoppaDSqlReplace.Tokens

Namespace Express

    ''' <summary>不等号式。</summary>
    Public NotInheritable Class NotEqualExpress
        Implements IExpression

        ' 左辺式
        Private ReadOnly mTml As IExpression

        ' 右辺式
        Private ReadOnly mTmr As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tml">左辺式。</param>
        ''' <param name="tmr">右辺式。</param>
        Public Sub New(tml As IExpression, tmr As IExpression)
            If tml IsNot Nothing AndAlso tmr IsNot Nothing Then
                Me.mTml = tml
                Me.mTmr = tmr
            Else
                Throw New DSqlAnalysisException("不等号式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Dim bval = Not EqualExpress.ExpressionEqual(Me.mTml?.Executes(env), Me.mTmr?.Executes(env))
            Return If(bval, CType(TrueToken.Value, IToken), FalseToken.Value)
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:<>"
        End Function

    End Class

End Namespace