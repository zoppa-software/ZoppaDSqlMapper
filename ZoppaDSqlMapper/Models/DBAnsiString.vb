Option Strict On
Option Explicit On

''' <summary>DbType.AnsiString型とマップさせるための文字列型です。</summary>
Public NotInheritable Class DBAnsiString
    Implements IComparable(Of DBAnsiString)

    ' 値
    Private ReadOnly mStr As String

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="str">格納する文字列。</param>
    Public Sub New(str As String)
        Me.mStr = str
    End Sub

    ''' <summary>文字列型にキャストします。</summary>
    ''' <param name="v">DBAnsiString型。</param>
    ''' <returns>文字列。</returns>
    Public Shared Widening Operator CType(ByVal v As DBAnsiString) As String
        Return v.mStr
    End Operator

    ''' <summary>DBAnsiString型にキャストします。</summary>
    ''' <param name="v">文字列。</param>
    ''' <returns>DBAnsiString型。</returns>
    Public Shared Widening Operator CType(ByVal v As String) As DBAnsiString
        Return New DBAnsiString(v)
    End Operator

    ''' <summary>文字列を取得します。</summary>
    ''' <returns>文字列。</returns>
    Public Overrides Function ToString() As String
        Return Me.mStr
    End Function

    ''' <summary>等しいか判定します。</summary>
    ''' <param name="obj">比較対象。</param>
    ''' <returns>比較結果。</returns>
    Public Overrides Function Equals(obj As Object) As Boolean
        Dim other = TryCast(obj, DBAnsiString)
        If other IsNot Nothing Then
            Return (Me.mStr = other.mStr)
        End If
        Return False
    End Function

    ''' <summary>ハッシュコード値を取得します。</summary>
    ''' <returns>ハッシュコード値。</returns>
    Public Overrides Function GetHashCode() As Integer
        Return Me.mStr.GetHashCode()
    End Function

    ''' <summary>比較を行います。</summary>
    ''' <param name="other">比較対象。</param>
    ''' <returns>比較結果。</returns>
    Public Function CompareTo(other As DBAnsiString) As Integer Implements IComparable(Of DBAnsiString).CompareTo
        Return Me.mStr.CompareTo(other.mStr)
    End Function

    ''' <summary>等号演算子です。</summary>
    ''' <param name="left">左辺値。</param>
    ''' <param name="right">右辺値。</param>
    ''' <returns>比較結果。</returns>
    Public Shared Operator =(left As DBAnsiString, right As DBAnsiString) As Boolean
        Return Not left.Equals(right)
    End Operator

    ''' <summary>不等号演算子です。</summary>
    ''' <param name="left">左辺値。</param>
    ''' <param name="right">右辺値。</param>
    ''' <returns>比較結果。</returns>
    Public Shared Operator <>(left As DBAnsiString, right As DBAnsiString) As Boolean
        Return left.Equals(right)
    End Operator

End Class
