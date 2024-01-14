Option Strict On
Option Explicit On

Imports ZoppaDSqlCompiler.Environments
Imports ZoppaDSqlCompiler.Tokens

Namespace Express

    ''' <summary>論理和式。</summary>
    Public NotInheritable Class OrExpress
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
                Throw New DSqlAnalysisException("論理和式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Dim tml = Me.mTml?.Executes(env)
            Dim tmr = Me.mTmr?.Executes(env)

            If TypeOf tml?.Contents Is Boolean AndAlso TypeOf tmr?.Contents Is Boolean Then
                Try
                    Dim lv = Convert.ToBoolean(tml.Contents)
                    Dim rv = Convert.ToBoolean(tmr.Contents)
                    Return If(lv OrElse rv, CType(TrueToken.Value, IToken), FalseToken.Value)
                Catch ex As Exception
                    Throw New DSqlAnalysisException($"論理和ができません。{tml.Contents} or {tmr.Contents}", ex)
                End Try
            Else
                Throw New DSqlAnalysisException($"論理和ができません。{tml.Contents} or {tmr.Contents}")
            End If
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:or"
        End Function

    End Class

End Namespace