Option Strict On
Option Explicit On

Imports System.Text

''' <summary>ユニークキーを表現するクラスです。</summary>
Public NotInheritable Class UniqueKey
    Implements IComparable(Of UniqueKey)

    ' キーリスト
    Private ReadOnly mKeys() As IComparable

    ''' <summary>コンストラクタ。</summary>
    ''' <param name="keys">キーリスト。</param>
    Private Sub New(keys As List(Of IComparable))
        Me.mKeys = keys.ToArray()
    End Sub

    ''' <summary>ユニークキーを作成します。</summary>
    ''' <param name="keys">キーリスト。</param>
    ''' <returns>ユニークキー。</returns>
    Public Shared Function Create(ParamArray keys As Object()) As UniqueKey
        Dim rkeys As New List(Of IComparable)()
        For Each k In keys
            Dim rk = TryCast(k, IComparable)
            If rk IsNot Nothing OrElse k Is Nothing Then
                rkeys.Add(rk)
            Else
                Throw New ArgumentException("キーが比較できません")
            End If
        Next
        If rkeys.Count > 0 Then
            Return New UniqueKey(rkeys)
        Else
            Throw New ArgumentException("キーが一つもありません")
        End If
    End Function

    ''' <summary>等しいか判定します。</summary>
    ''' <param name="obj">比較対象。</param>
    ''' <returns>等しければ真。</returns>
    Public Overrides Function Equals(obj As Object) As Boolean
        Dim other = TryCast(obj, UniqueKey)
        If other IsNot Nothing AndAlso
           Me.mKeys.Length = other.mKeys.Length Then
            Return (Me.CompareTo(other) = 0)
        Else
            Return False
        End If
    End Function

    ''' <summary>ユニークキーを比較します。</summary>
    ''' <param name="other">比較対象。</param>
    ''' <returns>比較結果。</returns>
    Public Function CompareTo(other As UniqueKey) As Integer Implements IComparable(Of UniqueKey).CompareTo
        For i As Integer = 0 To Math.Max(Me.mKeys.Length, other.mKeys.Length) - 1
            Dim lv = If(i < Me.mKeys.Length, Me.mKeys(i), Nothing)
            Dim rv = If(i < other.mKeys.Length, other.mKeys(i), Nothing)
            If lv Is Nothing AndAlso rv Is Nothing Then
            ElseIf lv Is Nothing Then
                Return -1
            ElseIf rv Is Nothing Then
                Return 1
            Else
                Dim res = lv.CompareTo(rv)
                If res <> 0 Then
                    Return res
                End If
            End If
        Next
        Return 0
    End Function

    ''' <summary>ハッシュ値を取得します。</summary>
    ''' <returns>ハッシュ値。</returns>
    Public Overrides Function GetHashCode() As Integer
        Dim res As Integer = If(Me.mKeys(0)?.GetHashCode(), 0)
        For i As Integer = 1 To Me.mKeys.Length - 1
            res = res Xor If(Me.mKeys(i)?.GetHashCode(), 0)
        Next
        Return res
    End Function

    ''' <summary>文字列表現を取得します。</summary>
    ''' <returns>文字列表現。</returns>
    Public Overrides Function ToString() As String
        Dim buf As New StringBuilder()
        For Each o In Me.mKeys
            If buf.Length > 0 Then buf.Append(", ")
            buf.Append(o)
        Next
        Return buf.ToString()
    End Function

End Class
