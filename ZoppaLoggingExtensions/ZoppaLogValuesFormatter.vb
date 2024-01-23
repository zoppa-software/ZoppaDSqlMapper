Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Text

Public NotInheritable Class ZoppaLogValuesFormatter

    Private Const NullValue As String = "(null)"

    Private Shared ReadOnly FormatDelimiters As Char() = New Char() {","c, ":"c}

    Private ReadOnly _valueNames As New List(Of String)()

    Private ReadOnly _format As String

    Public ReadOnly Property OriginalFormat As String

    Public ReadOnly Property ValueNames As List(Of String)
        Get
            Return Me._valueNames
        End Get
    End Property

    Public Sub New(format As String)
        If format Is Nothing Then
            Throw New NullReferenceException("有効な書式を指定していません")
        End If

        Me.OriginalFormat = format

        Dim vsb = New StringBuilder()
        Dim scanIndex As Integer = 0
        Dim endIndex As Integer = format.Length

        Do While scanIndex < endIndex
            Dim openBraceIndex As Integer = FindBraceIndex(format, "{"c, scanIndex, endIndex)
            If scanIndex = 0 AndAlso openBraceIndex = endIndex Then
                ' No holes found.
                Me._format = format
                Return
            End If

            Dim closeBraceIndex = FindBraceIndex(format, "}"c, openBraceIndex, endIndex)

            If closeBraceIndex = endIndex Then
                vsb.Append(format.Substring(scanIndex, endIndex - scanIndex))
                scanIndex = endIndex
            Else
                ' Format item syntax : { index[,alignment][ :formatString] }.
                Dim formatDelimiterIndex = FindIndexOfAny(format, FormatDelimiters, openBraceIndex, closeBraceIndex)

                vsb.Append(format.Substring(scanIndex, openBraceIndex - scanIndex + 1))
                vsb.Append(_valueNames.Count.ToString())
                _valueNames.Add(format.Substring(openBraceIndex + 1, formatDelimiterIndex - openBraceIndex - 1))
                vsb.Append(format.Substring(formatDelimiterIndex, closeBraceIndex - formatDelimiterIndex + 1))

                scanIndex = closeBraceIndex + 1
            End If

            Me._format = vsb.ToString()
        Loop
    End Sub

    Private Shared Function FindBraceIndex(format As String, brace As Char, startIndex As Integer, endIndex As Integer) As Integer
        ' Example: {{prefix{{{Argument}}}suffix}}.
        Dim braceIndex = endIndex
        Dim scanIndex = startIndex
        Dim braceOccurrenceCount = 0

        Do While scanIndex < endIndex
            If braceOccurrenceCount > 0 AndAlso format(scanIndex) <> brace Then
                If braceOccurrenceCount Mod 2 = 0 Then
                    ' Even number of '{' or '}' found. Proceed search with next occurrence of '{' or '}'.
                    braceOccurrenceCount = 0
                    braceIndex = endIndex
                Else
                    ' An unescaped '{' or '}' found.
                    Exit Do
                End If

            ElseIf format(scanIndex) = brace Then
                If brace = "}"c Then
                    If braceOccurrenceCount = 0 Then
                        ' For '}' pick the first occurrence.
                        braceIndex = scanIndex
                    End If
                Else
                    ' For '{' pick the last occurrence.
                    braceIndex = scanIndex
                End If

                braceOccurrenceCount += 1
            End If

            scanIndex += 1
        Loop

        Return braceIndex
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="format"></param>
    ''' <param name="chars">デリミタリスト。</param>
    ''' <param name="startIndex"></param>
    ''' <param name="endIndex"></param>
    ''' <returns></returns>
    Private Shared Function FindIndexOfAny(format As String, chars As Char(), startIndex As Integer, endIndex As Integer) As Integer
        Dim findIndex = format.IndexOfAny(chars, startIndex, endIndex - startIndex)
        Return If(findIndex = -1, endIndex, findIndex)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="values"></param>
    ''' <returns></returns>
    Public Function Format(values As Object()) As String
        Dim formattedValues = values

        If values IsNot Nothing Then
            For i As Integer = 0 To values.Length - 1
                Dim formattedValue = FormatArgument(values(i))

                ' If the formatted value is changed, we allocate and copy items to a new array to avoid mutating the array passed in to this method
                If Not ReferenceEquals(formattedValue, values(i)) Then
                    formattedValues = New Object(values.Length - 1) {}
                    Array.Copy(values, formattedValues, i)
                    formattedValues(i) = formattedValue
                    i += 1
                    Do While i < values.Length
                        formattedValues(i) = FormatArgument(values(i))
                        i += 1
                    Loop

                    Exit For
                End If
            Next
        End If

        Return String.Format(CultureInfo.InvariantCulture, Me._format, If(formattedValues, Array.Empty(Of Object)()))
    End Function

    ''' <summary>引数の配列のインデックス位置の要素を取得する。</summary>
    ''' <param name="values">配列。</param>
    ''' <param name="index">インデックス位置。</param>
    ''' <returns>要素。</returns>
    Public Function GetValue(values As Object(), index As Integer) As KeyValuePair(Of String, Object)
        ' 配列の範囲外ならエラー
        If index < 0 OrElse index > Me._valueNames.Count Then
            Throw New IndexOutOfRangeException(NameOf(index))
        End If

        ' 最後の要素なら元のメッセージを返す
        If Me._valueNames.Count > index Then
            Return New KeyValuePair(Of String, Object)(Me._valueNames(index), values(index))
        End If
        Return New KeyValuePair(Of String, Object)("{OriginalFormat}", OriginalFormat)
    End Function

    ''' <summary>ログの引数の文字列表現を取得する。</summary>
    ''' <param name="value">ログの引数</param>
    ''' <returns>変換した値。</returns>
    Private Shared Function FormatArgument(value As Object) As Object
        Dim stringValue As Object = Nothing
        Return If(TryFormatArgumentIfNullOrEnumerable(value, stringValue), stringValue, value)
    End Function

    ''' <summary>ログの引数の文字列表現を取得する。</summary>
    ''' <typeparam name="T">対象の型。</typeparam>
    ''' <param name="value">引数。</param>
    ''' <param name="stringValue">戻り値の文字列。</param>
    ''' <returns>変換したら真。</returns>
    Private Shared Function TryFormatArgumentIfNullOrEnumerable(Of T)(value As T, ByRef stringValue As Object) As Boolean
        ' 値が Nullなら Null文字列を返す
        If value Is Nothing Then
            stringValue = NullValue
            Return True
        End If

        ' 値が IEnumerableを実装しているが、それ自体が文字列でない場合はコンマ区切りの文字列を作成
        If TypeOf value IsNot String AndAlso TypeOf value Is IEnumerable Then
            Dim vsb = New StringBuilder(256)
            Dim first = True
            For Each e In DirectCast(value, IEnumerable)
                If Not first Then
                    vsb.Append(", ")
                End If

                vsb.Append(If(e IsNot Nothing, e.ToString(), NullValue))
                first = False
            Next

            stringValue = vsb.ToString()
            Return True
        End If

        Return False
    End Function

End Class
