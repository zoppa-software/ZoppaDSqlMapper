Option Strict On
Option Explicit On
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports ZoppaDSqlReplace.Tokens.NumberToken

Namespace Tokens

    ''' <summary>数値トークン。</summary>
    Public MustInherit Class NumberToken
        Implements IToken

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public MustOverride ReadOnly Property Contents As Object Implements IToken.Contents

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        Public ReadOnly Property TokenType As Type Implements IToken.TokenType
            Get
                Return GetType(NumberToken)
            End Get
        End Property

        ''' <summary>トークンが空白文字ならば真を返します。</summary>
        ''' <returns>トークンが空白文字ならば真。</returns>
        Public ReadOnly Property IsWhiteSpace As Boolean Implements IToken.IsWhiteSpace
            Get
                Return False
            End Get
        End Property

        ''' <summary>トークンが改行文字ならば真を返します。</summary>
        ''' <returns>トークンが改行文字ならば真。</returns>
        Public ReadOnly Property IsCrLf As Boolean Implements IToken.IsCrLf
            Get
                Return False
            End Get
        End Property

        ''' <summary>文字列を数値トークンに変換します。</summary>
        ''' <param name="input">入力文字列。</param>
        ''' <returns>数値トークン。</returns>
        Friend Shared Function Create(input As String) As NumberToken
            Dim dval As Double = 0
            If Double.TryParse(input, dval) Then
                Return NumberToken.ConvertToken(dval)
            Else
                Throw New DSqlAnalysisException("文字列を数値トークンに変換できません")
            End If
        End Function

        ''' <summary>数値トークンを加算します。</summary>
        ''' <param name="other">加算する値。</param>
        ''' <returns>加算結果。</returns>
        Public MustOverride Function PlusComputation(other As NumberToken) As NumberToken

        ''' <summary>数値トークンを減算します。</summary>
        ''' <param name="other">減算する値。</param>
        ''' <returns>減算結果。</returns>
        Public MustOverride Function MinusComputation(other As NumberToken) As NumberToken

        ''' <summary>数値トークンを乗算します。</summary>
        ''' <param name="other">乗算する値。</param>
        ''' <returns>乗算結果。</returns>
        Public MustOverride Function MultiComputation(other As NumberToken) As NumberToken

        ''' <summary>数値トークンを除算します。</summary>
        ''' <param name="other">除算する値。</param>
        ''' <returns>除算結果。</returns>
        Public MustOverride Function DivComputation(other As NumberToken) As NumberToken

        ''' <summary>数値が等しければ真を返します。</summary>
        ''' <param name="other">比較する数値トークン。</param>
        ''' <returns>等しければ真。</returns>
        Public MustOverride Function EqualCondition(other As NumberToken) As Boolean

        ''' <summary>数値トークンの符号を逆転します。</summary>
        ''' <returns>数値トークン。</returns>
        Public MustOverride Function SignChange() As NumberToken

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return Me.Contents.ToString()
        End Function

        ''' <summary>オブジェクトを数値トークンへ変換します。</summary>
        ''' <param name="obj">オブジェクト。</param>
        ''' <returns>数値トークン。</returns>
        Public Shared Function ConvertToken(obj As Object) As NumberToken
            Dim ans As NumberToken = Nothing

            If TypeOf obj Is ULong Then
                Throw New InvalidCastException("符号なし Int64型はサポート外です")
            ElseIf TypeOf obj Is Double OrElse TypeOf obj Is Single Then
                Dim dval = Convert.ToDouble(obj)
                ans = New NumberDoubleValueToken(dval)
                If Math.Truncate(dval) <> dval Then
                    Return ans
                End If
            ElseIf TypeOf obj Is Decimal Then
                Dim dval = Convert.ToDecimal(obj)
                ans = New NumberDecimalValueToken(dval)
                If Math.Truncate(dval) <> dval Then
                    Return ans
                End If
            End If

            ' Integerに変換可能ならば Integerにする
            Try
                Return New NumberIntegerValueToken(Convert.ToInt32(obj))
            Catch ex As OverflowException

            End Try
            ' Longに変換可能ならば Longにする
            Try
                Return New NumberLongValueToken(Convert.ToInt64(obj))
            Catch ex As OverflowException

            End Try
            Return ans
        End Function

        ''' <summary>数値トークン（Integer）</summary>
        Public NotInheritable Class NumberIntegerValueToken
            Inherits NumberToken

            ' 数値
            Private ReadOnly mValue As Integer

            ''' <summary>格納されている値を取得する。</summary>
            ''' <returns>格納値。</returns>
            Public Overrides ReadOnly Property Contents As Object
                Get
                    Return Me.mValue
                End Get
            End Property

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="value">数値。</param>
            Public Sub New(value As Integer)
                Me.mValue = value
            End Sub

            ''' <summary>数値トークンを加算します。</summary>
            ''' <param name="other">加算する値。</param>
            ''' <returns>加算結果。</returns>
            Public Overrides Function PlusComputation(other As NumberToken) As NumberToken
                Dim ans As Object
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    ans = self + oval
                ElseIf TypeOf other.Contents Is Decimal Then
                    Dim self = Convert.ToDecimal(Me.Contents)
                    Dim oval = Convert.ToDecimal(other.Contents)
                    ans = self + oval
                ElseIf TypeOf other.Contents Is Long Then
                    Dim self = Convert.ToInt64(Me.Contents)
                    Dim oval = Convert.ToInt64(other.Contents)
                    ans = self + oval
                Else
                    Dim oval = Convert.ToInt32(other.Contents)
                    ans = Me.mValue + oval
                End If
                Return NumberToken.ConvertToken(ans)
            End Function

            ''' <summary>数値トークンを減算します。</summary>
            ''' <param name="other">減算する値。</param>
            ''' <returns>減算結果。</returns>
            Public Overrides Function MinusComputation(other As NumberToken) As NumberToken
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    Return New NumberDoubleValueToken(self - oval)
                ElseIf TypeOf other.Contents Is Decimal Then
                    Dim self = Convert.ToDecimal(Me.Contents)
                    Dim oval = Convert.ToDecimal(other.Contents)
                    Return New NumberDecimalValueToken(self - oval)
                ElseIf TypeOf other.Contents Is Long Then
                    Dim self = Convert.ToInt64(Me.Contents)
                    Dim oval = Convert.ToInt64(other.Contents)
                    Return New NumberLongValueToken(self - oval)
                Else
                    Dim oval = Convert.ToInt32(other.Contents)
                    Return New NumberIntegerValueToken(Me.mValue - oval)
                End If
            End Function

            ''' <summary>数値トークンを乗算します。</summary>
            ''' <param name="other">乗算する値。</param>
            ''' <returns>乗算結果。</returns>
            Public Overrides Function MultiComputation(other As NumberToken) As NumberToken
                Dim ans As Object
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    ans = self * oval
                ElseIf TypeOf other.Contents Is Decimal Then
                    Dim self = Convert.ToDecimal(Me.Contents)
                    Dim oval = Convert.ToDecimal(other.Contents)
                    ans = self * oval
                ElseIf TypeOf other.Contents Is Long Then
                    Dim self = Convert.ToInt64(Me.Contents)
                    Dim oval = Convert.ToInt64(other.Contents)
                    ans = self * oval
                Else
                    Dim oval = Convert.ToInt32(other.Contents)
                    ans = Me.mValue * oval
                End If
                Return NumberToken.ConvertToken(ans)
            End Function

            ''' <summary>数値トークンを除算します。</summary>
            ''' <param name="other">除算する値。</param>
            ''' <returns>除算結果。</returns>
            Public Overrides Function DivComputation(other As NumberToken) As NumberToken
                Dim ans As Object
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    If Math.Abs(oval) < Double.Epsilon * 2 Then
                        Throw New DivideByZeroException("0割が発生しました")
                    End If
                    ans = self / oval
                ElseIf TypeOf other.Contents Is Decimal Then
                    Dim self = Convert.ToDecimal(Me.Contents)
                    Dim oval = Convert.ToDecimal(other.Contents)
                    If Math.Abs(oval) < Double.Epsilon * 2 Then
                        Throw New DivideByZeroException("0割が発生しました")
                    End If
                    ans = self / oval
                ElseIf TypeOf other.Contents Is Long Then
                    Dim self = Convert.ToInt64(Me.Contents)
                    Dim oval = Convert.ToInt64(other.Contents)
                    If Math.Abs(oval) < Double.Epsilon * 2 Then
                        Throw New DivideByZeroException("0割が発生しました")
                    End If
                    ans = self / oval
                Else
                    Dim oval = Convert.ToInt32(other.Contents)
                    If Math.Abs(oval) < Double.Epsilon * 2 Then
                        Throw New DivideByZeroException("0割が発生しました")
                    End If
                    ans = Me.mValue / oval
                End If
                Return NumberToken.ConvertToken(ans)
            End Function

            ''' <summary>数値が等しければ真を返します。</summary>
            ''' <param name="other">比較する数値トークン。</param>
            ''' <returns>等しければ真。</returns>
            Public Overrides Function EqualCondition(other As NumberToken) As Boolean
                If TypeOf other.Contents Is Double Then
                    Return other.EqualCondition(Me)
                Else
                    Return Me.mValue.Equals(other.Contents)
                End If
            End Function

            ''' <summary>数値トークンの符号を逆転します。</summary>
            ''' <returns>数値トークン。</returns>
            Public Overrides Function SignChange() As NumberToken
                Return New NumberIntegerValueToken(-Me.mValue)
            End Function

        End Class

        ''' <summary>数値トークン（Long）</summary>
        Public NotInheritable Class NumberLongValueToken
            Inherits NumberToken

            ' 数値
            Private ReadOnly mValue As Long

            ''' <summary>格納されている値を取得する。</summary>
            ''' <returns>格納値。</returns>
            Public Overrides ReadOnly Property Contents As Object
                Get
                    Return Me.mValue
                End Get
            End Property

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="value">数値。</param>
            Public Sub New(value As Long)
                Me.mValue = value
            End Sub

            ''' <summary>数値トークンを加算します。</summary>
            ''' <param name="other">加算する値。</param>
            ''' <returns>加算結果。</returns>
            Public Overrides Function PlusComputation(other As NumberToken) As NumberToken
                Dim ans As Object
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    ans = self + oval
                ElseIf TypeOf other.Contents Is Decimal Then
                    Dim self = Convert.ToDecimal(Me.Contents)
                    Dim oval = Convert.ToDecimal(other.Contents)
                    ans = self + oval
                Else
                    Dim oval = Convert.ToInt64(other.Contents)
                    ans = Me.mValue + oval
                End If
                Return NumberToken.ConvertToken(ans)
            End Function

            ''' <summary>数値トークンを減算します。</summary>
            ''' <param name="other">減算する値。</param>
            ''' <returns>減算結果。</returns>
            Public Overrides Function MinusComputation(other As NumberToken) As NumberToken
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    Return New NumberDoubleValueToken(self - oval)
                ElseIf TypeOf other.Contents Is Decimal Then
                    Dim self = Convert.ToDecimal(Me.Contents)
                    Dim oval = Convert.ToDecimal(other.Contents)
                    Return New NumberDecimalValueToken(self - oval)
                Else
                    Dim oval = Convert.ToInt64(other.Contents)
                    Return New NumberLongValueToken(Me.mValue - oval)
                End If
            End Function

            ''' <summary>数値トークンを乗算します。</summary>
            ''' <param name="other">乗算する値。</param>
            ''' <returns>乗算結果。</returns>
            Public Overrides Function MultiComputation(other As NumberToken) As NumberToken
                Dim ans As Object
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    ans = self * oval
                ElseIf TypeOf other.Contents Is Decimal Then
                    Dim self = Convert.ToDecimal(Me.Contents)
                    Dim oval = Convert.ToDecimal(other.Contents)
                    ans = self * oval
                Else
                    Dim oval = Convert.ToInt64(other.Contents)
                    ans = Me.mValue * oval
                End If
                Return NumberToken.ConvertToken(ans)
            End Function

            ''' <summary>数値トークンを除算します。</summary>
            ''' <param name="other">除算する値。</param>
            ''' <returns>除算結果。</returns>
            Public Overrides Function DivComputation(other As NumberToken) As NumberToken
                Dim ans As Object
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    If Math.Abs(oval) < Double.Epsilon * 2 Then
                        Throw New DivideByZeroException("0割が発生しました")
                    End If
                    ans = self / oval
                ElseIf TypeOf other.Contents Is Decimal Then
                    Dim self = Convert.ToDecimal(Me.Contents)
                    Dim oval = Convert.ToDecimal(other.Contents)
                    If Math.Abs(oval) < Double.Epsilon * 2 Then
                        Throw New DivideByZeroException("0割が発生しました")
                    End If
                    ans = self / oval
                Else
                    Dim oval = Convert.ToInt64(other.Contents)
                    If Math.Abs(oval) < Double.Epsilon * 2 Then
                        Throw New DivideByZeroException("0割が発生しました")
                    End If
                    ans = Me.mValue / oval
                End If
                Return NumberToken.ConvertToken(ans)
            End Function

            ''' <summary>数値が等しければ真を返します。</summary>
            ''' <param name="other">比較する数値トークン。</param>
            ''' <returns>等しければ真。</returns>
            Public Overrides Function EqualCondition(other As NumberToken) As Boolean
                If TypeOf other.Contents Is Double Then
                    Return other.EqualCondition(Me)
                Else
                    Return Me.mValue.Equals(other.Contents)
                End If
            End Function

            ''' <summary>数値トークンの符号を逆転します。</summary>
            ''' <returns>数値トークン。</returns>
            Public Overrides Function SignChange() As NumberToken
                Return New NumberLongValueToken(-Me.mValue)
            End Function

        End Class

        ''' <summary>数値トークン（Decimal）</summary>
        Public NotInheritable Class NumberDecimalValueToken
            Inherits NumberToken

            ' 数値
            Private ReadOnly mValue As Decimal

            ''' <summary>格納されている値を取得する。</summary>
            ''' <returns>格納値。</returns>
            Public Overrides ReadOnly Property Contents As Object
                Get
                    Return Me.mValue
                End Get
            End Property

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="value">数値。</param>
            Public Sub New(value As Decimal)
                Me.mValue = value
            End Sub

            ''' <summary>数値トークンを加算します。</summary>
            ''' <param name="other">加算する値。</param>
            ''' <returns>加算結果。</returns>
            Public Overrides Function PlusComputation(other As NumberToken) As NumberToken
                Dim ans As Object
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    ans = self + oval
                Else
                    Dim oval = Convert.ToDecimal(other.Contents)
                    ans = Me.mValue + oval
                End If
                Return NumberToken.ConvertToken(ans)
            End Function

            ''' <summary>数値トークンを減算します。</summary>
            ''' <param name="other">減算する値。</param>
            ''' <returns>減算結果。</returns>
            Public Overrides Function MinusComputation(other As NumberToken) As NumberToken
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    Return New NumberDoubleValueToken(self - oval)
                Else
                    Dim oval = Convert.ToDecimal(other.Contents)
                    Return New NumberDecimalValueToken(Me.mValue - oval)
                End If
            End Function

            ''' <summary>数値トークンを乗算します。</summary>
            ''' <param name="other">乗算する値。</param>
            ''' <returns>乗算結果。</returns>
            Public Overrides Function MultiComputation(other As NumberToken) As NumberToken
                Dim ans As Object
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    ans = self * oval
                Else
                    Dim oval = Convert.ToDecimal(other.Contents)
                    ans = Me.mValue * oval
                End If
                Return NumberToken.ConvertToken(ans)
            End Function

            ''' <summary>数値トークンを除算します。</summary>
            ''' <param name="other">除算する値。</param>
            ''' <returns>除算結果。</returns>
            Public Overrides Function DivComputation(other As NumberToken) As NumberToken
                Dim ans As Object
                If TypeOf other.Contents Is Double Then
                    Dim self = Convert.ToDouble(Me.Contents)
                    Dim oval = Convert.ToDouble(other.Contents)
                    If Math.Abs(oval) < Double.Epsilon * 2 Then
                        Throw New DivideByZeroException("0割が発生しました")
                    End If
                    ans = self / oval
                Else
                    Dim oval = Convert.ToDecimal(other.Contents)
                    If Math.Abs(oval) < Double.Epsilon * 2 Then
                        Throw New DivideByZeroException("0割が発生しました")
                    End If
                    ans = Me.mValue / oval
                End If
                Return NumberToken.ConvertToken(ans)
            End Function

            ''' <summary>数値が等しければ真を返します。</summary>
            ''' <param name="other">比較する数値トークン。</param>
            ''' <returns>等しければ真。</returns>
            Public Overrides Function EqualCondition(other As NumberToken) As Boolean
                If TypeOf other.Contents Is Double Then
                    Return other.EqualCondition(Me)
                Else
                    Return Me.mValue.Equals(other.Contents)
                End If
            End Function

            ''' <summary>数値トークンの符号を逆転します。</summary>
            ''' <returns>数値トークン。</returns>
            Public Overrides Function SignChange() As NumberToken
                Return New NumberDecimalValueToken(-Me.mValue)
            End Function

        End Class

        ''' <summary>数値トークン（Double）</summary>
        Public NotInheritable Class NumberDoubleValueToken
            Inherits NumberToken

            ' 数値
            Private ReadOnly mValue As Double

            ''' <summary>格納されている値を取得する。</summary>
            ''' <returns>格納値。</returns>
            Public Overrides ReadOnly Property Contents As Object
                Get
                    Return Me.mValue
                End Get
            End Property

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="value">数値。</param>
            Public Sub New(value As Double)
                Me.mValue = value
            End Sub

            ''' <summary>数値トークンを加算します。</summary>
            ''' <param name="other">加算する値。</param>
            ''' <returns>加算結果。</returns>
            Public Overrides Function PlusComputation(other As NumberToken) As NumberToken
                Dim oval = Convert.ToDouble(other.Contents)
                Return New NumberDoubleValueToken(Me.mValue + oval)
            End Function

            ''' <summary>数値トークンを減算します。</summary>
            ''' <param name="other">減算する値。</param>
            ''' <returns>減算結果。</returns>
            Public Overrides Function MinusComputation(other As NumberToken) As NumberToken
                Dim oval = Convert.ToDouble(other.Contents)
                Return New NumberDoubleValueToken(Me.mValue - oval)
            End Function

            ''' <summary>数値トークンを乗算します。</summary>
            ''' <param name="other">乗算する値。</param>
            ''' <returns>乗算結果。</returns>
            Public Overrides Function MultiComputation(other As NumberToken) As NumberToken
                Dim oval = Convert.ToDouble(other.Contents)
                Return New NumberDoubleValueToken(Me.mValue * oval)
            End Function

            ''' <summary>数値トークンを除算します。</summary>
            ''' <param name="other">除算する値。</param>
            ''' <returns>除算結果。</returns>
            Public Overrides Function DivComputation(other As NumberToken) As NumberToken
                Dim oval = Convert.ToDouble(other.Contents)
                If Math.Abs(oval) < Double.Epsilon * 2 Then
                    Throw New DivideByZeroException("0割が発生しました")
                End If
                Return New NumberDoubleValueToken(Me.mValue / oval)
            End Function

            ''' <summary>数値が等しければ真を返します。</summary>
            ''' <param name="other">比較する数値トークン。</param>
            ''' <returns>等しければ真。</returns>
            Public Overrides Function EqualCondition(other As NumberToken) As Boolean
                Dim oval = Convert.ToDouble(other.Contents)
                Return (Math.Abs(Me.mValue - oval) < 0.00000000000001)
            End Function

            ''' <summary>数値トークンの符号を逆転します。</summary>
            ''' <returns>数値トークン。</returns>
            Public Overrides Function SignChange() As NumberToken
                Return New NumberDoubleValueToken(-Me.mValue)
            End Function

        End Class

    End Class

End Namespace
