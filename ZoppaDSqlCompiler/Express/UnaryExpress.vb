Option Strict On
Option Explicit On

Imports ZoppaDSqlCompiler.Environments
Imports ZoppaDSqlCompiler.Tokens

Namespace Express

    ''' <summary>前置き式。</summary>
    Public NotInheritable Class UnaryExpress
        Implements IExpression

        ' 対象トークン
        Private ReadOnly mToken As IToken

        ' 対象式
        Private ReadOnly mValue As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="token">対象トークン。</param>
        Public Sub New(token As IToken, factor As IExpression)
            If token IsNot Nothing AndAlso factor IsNot Nothing Then
                Me.mToken = token
                Me.mValue = factor
            Else
                Throw New DSqlAnalysisException("前置き式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Dim tkn = Me.mValue?.Executes(env)

            Select Case Me.mToken?.TokenName
                Case NameOf(PlusToken)
                    If tkn?.TokenName = NameOf(NumberToken) Then
                        Return tkn
                    Else
                        Throw New DSqlAnalysisException("前置き+が数字の前に置かれていません")
                    End If

                Case NameOf(MinusToken)
                    If tkn?.TokenName = NameOf(NumberToken) Then
                        Return CType(tkn, NumberToken).SignChange()
                    Else
                        Throw New DSqlAnalysisException("前置き-が数字の前に置かれていません")
                    End If

                Case NameOf(NotToken)
                    Dim bval = Convert.ToBoolean(tkn?.Contents)
                    Return If(bval, CType(FalseToken.Value, IToken), TrueToken.Value)

                Case Else
                    Throw New DSqlAnalysisException("有効な前置き式ではありません")
            End Select
        End Function
    End Class

End Namespace
