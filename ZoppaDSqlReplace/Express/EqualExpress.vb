Option Strict On
Option Explicit On

Imports ZoppaDSqlReplace.Environments
Imports ZoppaDSqlReplace.Tokens

Namespace Express

    ''' <summary>等号式。</summary>
    Public NotInheritable Class EqualExpress
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
                Throw New DSqlAnalysisException("等号式の生成にNullは使用できません")
            End If
        End Sub

        ''' <summary>式を実行する。</summary>
        ''' <param name="env">環境値情報。</param>
        ''' <returns>実行結果。</returns>
        Public Function Executes(env As IEnvironmentValue) As IToken Implements IExpression.Executes
            Dim bval = ExpressionEqual(Me.mTml?.Executes(env), Me.mTmr?.Executes(env))
            Return If(bval, CType(TrueToken.Value, IToken), FalseToken.Value)
        End Function

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "expr:="
        End Function

        ''' <summary>式の等価比較を行います。</summary>
        ''' <param name="tml">左辺トークン。</param>
        ''' <param name="tmr">右辺トークン。</param>
        ''' <returns>等価比較結果。</returns>
        Friend Shared Function ExpressionEqual(tml As IToken, tmr As IToken) As Boolean
            Dim nml = TryCast(tml, NumberToken)
            Dim nmr = TryCast(tmr, NumberToken)

            If nml IsNot Nothing AndAlso nmr IsNot Nothing Then
                ' 数値比較
                Return nml.EqualCondition(nmr)
            ElseIf tml?.Contents Is Nothing AndAlso tmr?.Contents Is Nothing Then
                ' 両方 Null比較
                Return True
            ElseIf TypeOf tml?.Contents Is [Enum] Then
                ' 列挙型比較(左辺)
                Return ExpressionEnumEqual(tml, tmr, nmr)
            ElseIf TypeOf tmr?.Contents Is [Enum] Then
                ' 列挙型比較(右辺)
                Return ExpressionEnumEqual(tmr, tml, nml)
            Else
                ' その他
                Return If(tml?.Contents?.Equals(tmr?.Contents), False)
            End If
        End Function

        ''' <summary>列挙型比較を行います。</summary>
        ''' <param name="enumToken">列挙型トークン。</param>
        ''' <param name="otherToken">その他トークン。</param>
        ''' <param name="numToken">数値トークン。</param>
        ''' <returns>比較結果。</returns>
        Private Shared Function ExpressionEnumEqual(enumToken As IToken, otherToken As IToken, numToken As IToken) As Boolean
            If numToken IsNot Nothing Then
                Return If(numToken.Contents?.Equals(CInt(enumToken.Contents)), False)
            ElseIf otherToken?.TokenType Is GetType(StringToken) Then
                Return If(otherToken.Contents?.Equals([Enum].GetName(enumToken.Contents.GetType(), enumToken.Contents)), False)
            Else
                Return If(otherToken.Contents?.Equals(enumToken.Contents), False)
            End If
        End Function

    End Class

End Namespace