Option Strict On
Option Explicit On

''' <summary>ユニークキー管理リスト。</summary>
''' <typeparam name="T">対象の型。</typeparam>
Public Class PrimaryKeyList(Of T)
    Implements IEnumerable(Of T)

    ' ユニークキーリスト
    Private ReadOnly mInnerList As New SortedDictionary(Of UniqueKey, T)

    ''' <summary>指定位置のユニークキーを取得します。</summary>
    ''' <param name="index">指定位置。</param>
    ''' <returns>ユニークキー。</returns>
    Default Public ReadOnly Property Item(index As Integer) As T
        Get
            If index >= 0 AndAlso index < Me.mInnerList.Keys.Count Then
                Dim key = Me.mInnerList.Keys(index)
                Return Me.mInnerList(key)
            Else
                Throw New IndexOutOfRangeException("インデックスが要素の範囲外です")
            End If
        End Get
    End Property

    ''' <summary>ユニークキーの数を取得します。</summary>
    Public ReadOnly Property Count As Integer
        Get
            Return Me.mInnerList.Count
        End Get
    End Property

    ''' <summary>指定位置のユニークキーを削除します。</summary>
    ''' <param name="index">指定位置。</param>
    Public Sub RemoveAt(index As Integer)
        If index >= 0 AndAlso index < Me.mInnerList.Keys.Count Then
            Dim key = Me.mInnerList.Keys(index)
            Me.mInnerList.Remove(key)
        Else
            Throw New IndexOutOfRangeException("インデックスが要素の範囲外です")
        End If
    End Sub

    ''' <summary>ユニークキーをキーを指定して登録します。</summary>
    ''' <param name="item">ユニークキー。</param>
    ''' <param name="primaryKey">キーリスト。</param>
    Public Sub Regist(item As T, ParamArray primaryKey As Object())
        Dim key = UniqueKey.Create(primaryKey)
        Me.mInnerList.Add(key, item)
    End Sub

    ''' <summary>ユニークキーをキーを指定して登録します。</summary>
    ''' <param name="primaryKey">キーリスト。</param>
    ''' <param name="item">ユニークキー。</param>
    Public Sub Regist(primaryKey As Object(), item As T)
        Dim key = UniqueKey.Create(primaryKey)
        Me.mInnerList.Add(key, item)
    End Sub

    ''' <summary>ユニークキーをキーを指定して登録します。</summary>
    ''' <param name="primaryKey">キーリスト。</param>
    ''' <param name="item">ユニークキー。</param>
    Public Sub Regist(primaryKey As UniqueKey, item As T)
        Me.mInnerList.Add(primaryKey, item)
    End Sub

    ''' <summary>ユニークキーを消去します。</summary>
    Public Sub Clear()
        Me.mInnerList.Clear()
    End Sub

    ''' <summary>指定したユニークキーの位置を取得します。</summary>
    ''' <param name="item">ユニークキー。</param>
    ''' <returns>位置。</returns>
    Public Function IndexOf(item As T) As Integer
        Dim i As Integer = 0
        For Each pair In Me.mInnerList
            If pair.Value.Equals(item) Then
                Return i
            End If
            i += 1
        Next
        Return -1
    End Function

    ''' <summary>指定したユニークキーが登録していたら真を返します。</summary>
    ''' <param name="item">ユニークキー。</param>
    ''' <returns>登録されていたら真。</returns>
    Public Function Contains(item As T) As Boolean
        For Each pair In Me.mInnerList
            If pair.Value.Equals(item) Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>指定したキーのユニークキーがあれば真を返します。</summary>
    ''' <param name="keys">キー。</param>
    ''' <returns>キーがあれば真。</returns>
    Public Function ContainsKey(ParamArray keys As Object()) As Boolean
        Dim key = UniqueKey.Create(keys)
        Return Me.mInnerList.ContainsKey(key)
    End Function

    ''' <summary>指定したキーのユニークキーを取得します。</summary>
    ''' <param name="keys">キー。</param>
    ''' <returns>ユニークキー。</returns>
    Public Function GetValue(ParamArray keys As Object()) As T
        Dim key = UniqueKey.Create(keys)
        Return Me.mInnerList(key)
    End Function

    ''' <summary>指定したキーのユニークキーがあれば真を返します。</summary>
    ''' <param name="v">ユニークキー（戻り値）</param>
    ''' <param name="keys">キー。</param>
    ''' <returns>登録があれば真を返します。</returns>
    Public Function TrySearchValue(ByRef v As T, ParamArray keys As Object()) As Boolean
        Dim key = UniqueKey.Create(keys)
        Return Me.mInnerList.TryGetValue(key, v)
    End Function

    ''' <summary>指定したキーのユニークキーがあれば真とユニークキーを返します。</summary>
    ''' <param name="keys">キー。</param>
    ''' <returns>登録の有無を表す真偽値とユニークキー。</returns>
    Public Function SearchValue(ParamArray keys As Object()) As (hasValue As Boolean, value As T)
        Dim key = UniqueKey.Create(keys)
        Dim v As T
        If Me.mInnerList.TryGetValue(key, v) Then
            Return (True, v)
        Else
            Return (False, Nothing)
        End If
    End Function

    ''' <summary>指定したユニークキーを削除します。</summary>
    ''' <param name="item">ユニークキー。</param>
    ''' <returns>削除できたら真。</returns>
    Public Function Remove(item As T) As Boolean
        Dim key As UniqueKey = Nothing
        For Each pair In Me.mInnerList
            If pair.Value.Equals(item) Then
                key = pair.Key
                Exit For
            End If
        Next
        If key IsNot Nothing Then
            Me.mInnerList.Remove(key)
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>列挙子を取得します。</summary>
    ''' <returns>列挙子。</returns>
    Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
        Return Me.mInnerList.Values.GetEnumerator()
    End Function

    ''' <summary>列挙子を取得します。</summary>
    ''' <returns>列挙子。</returns>
    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

End Class
