Option Strict On
Option Explicit On

Imports ZoppaDSqlReplace.Environments
Imports ZoppaDSqlReplace.Tokens

Namespace Express

    ''' <summary>メソッド式。</summary>
    Public NotInheritable Class MethodEvalExpress
        Implements IExpression

        ' メソッド名
        Private ReadOnly mMethod As IdentToken

        ' 真式
        Private ReadOnly mArgs As IExpression

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="name">条件式。</param>
        ''' <param name="args">真式。</param>
        Public Sub New(name As IdentToken, args As IExpression)
            If name IsNot Nothing AndAlso args IsNot Nothing Then
                Me.mMethod = name
                Me.mArgs = args
            Else
                Throw New DSqlAnalysisException("メソッド式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Dim mtd = EnvironmentObjectValue.GetMethod(If(Me.mMethod.Contents?.ToString(), ""))
            Dim prm = CType(Me.mArgs.Executes(env).Contents, Object())
            Return mtd.Invoke(prm)
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return $"expr:method {Me.mMethod}"
        End Function

    End Class

End Namespace

