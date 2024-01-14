Option Strict On
Option Explicit On

Imports ZoppaDSqlCompiler.Environments
Imports ZoppaDSqlCompiler.Tokens

Namespace Express

    ''' <summary>括弧内部式。</summary>
    Public NotInheritable Class ParenExpress
        Implements IExpression

        ' 内部式
        Private mInner As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inner">内部式。</param>
        Public Sub New(inner As IExpression)
            Me.mInner = inner
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Dim ans = Me.mInner?.Executes(env)
            If ans IsNot Nothing Then
                Return ans
            Else
                Throw New DSqlAnalysisException("内部式が実行できません")
            End If
        End Function

    End Class

End Namespace