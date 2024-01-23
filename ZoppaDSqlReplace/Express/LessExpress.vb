Option Strict On
Option Explicit On

Imports ZoppaDSqlReplace.Environments
Imports ZoppaDSqlReplace.Tokens

Namespace Express

    ''' <summary>小なり式。</summary>
    Public NotInheritable Class LessExpress
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
                Throw New DSqlAnalysisException("小なり式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Dim tml = Me.mTml?.Executes(env)
            Dim tmr = Me.mTmr?.Executes(env)

            ' 左右辺の値が誤差未満なら等しいためFalse
            Dim nml = TryCast(tml, NumberToken)
            Dim nmr = TryCast(tmr, NumberToken)
            If nml IsNot Nothing AndAlso nmr IsNot Nothing Then
                If nml.EqualCondition(nmr) Then
                    Return FalseToken.Value
                End If
            End If

            ' 比較インタフェースを取得して結果を取得
            Dim comp = TryCast(tml?.Contents, IComparable)
            If comp IsNot Nothing Then
                Dim bval = (comp.CompareTo(tmr?.Contents) < 0)
                Return If(bval, CType(TrueToken.Value, IToken), FalseToken.Value)
            Else
                Throw New DSqlAnalysisException($"比較ができません。{tml.Contents} < {tmr.Contents}")
            End If
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:<"
        End Function

    End Class

End Namespace