Option Strict On
Option Explicit On

Imports System.Text
Imports ZoppaDSqlCompiler.Tokens
Imports ZoppaDSqlCompiler.TokenCollection

''' <summary>字句解析機能。</summary>
Friend Module LexicalAnalysis

    ''' <summary>文字列をトークン解析します。</summary>
    ''' <param name="command">文字列。</param>
    ''' <returns>トークンリスト。</returns>
    Public Function Compile(command As String) As List(Of TokenPosition)
        Dim tokens As New List(Of TokenPosition)()

        Dim reader = New StringPtr(command)

        Dim buffer As New StringBuilder()
        Dim token As IToken
        Dim pos As Integer
        Do While reader.HasNext
            Dim c = reader.CurrentToMove()
            Select Case c
                Case "{"c
                    pos = reader.CurrentPosition
                    token = CreateQueryToken(buffer)
                    If token IsNot Nothing Then tokens.Add(New TokenPosition(token, pos))

                    pos = reader.CurrentPosition
                    token = CreateTokens(reader)
                    If token IsNot Nothing Then tokens.Add(New TokenPosition(token, pos))

                Case "#"c, "!"c, "$"c
                    If reader.Current = "{"c Then
                        pos = reader.CurrentPosition
                        token = CreateQueryToken(buffer)
                        If token IsNot Nothing Then tokens.Add(New TokenPosition(token, pos))

                        reader.CurrentToMove()
                        pos = reader.CurrentPosition
                        token = CreateReplaseToken(reader, c = "#"c)
                        tokens.Add(New TokenPosition(token, pos))
                    Else
                        buffer.Append(c)
                    End If

                Case "\"c
                    If reader.Current = "{"c OrElse reader.Current = "}"c Then
                        buffer.Append(reader.Current)
                        reader.CurrentToMove()
                    End If

                Case Else
                    buffer.Append(c)
            End Select
        Loop

        pos = reader.CurrentPosition
        token = CreateQueryToken(buffer)
        If token IsNot Nothing Then tokens.Add(New TokenPosition(token, pos))

        Return tokens
    End Function

    ''' <summary>直接出力するクエリトークンを生成します。</summary>
    ''' <param name="buffer">文字列バッファ。</param>
    ''' <returns>クエリトークン。</returns>
    Private Function CreateQueryToken(buffer As StringBuilder) As IToken
        Dim res As IToken = Nothing
        If buffer.Length > 0 Then
            res = New QueryToken(buffer.ToString())
            buffer.Clear()
        End If
        Return res
    End Function

    ''' <summary>置き換えトークンを生成します。</summary>
    ''' <param name="reader">入力文字ストリーム。</param>
    ''' <param name="isEscape">エスケープするならば真。</param>
    ''' <returns>置き換えトークン。</returns>
    Private Function CreateReplaseToken(reader As StringPtr, isEscape As Boolean) As IToken
        Return New ReplaseToken(GetPartitionString(reader, "置き換えトークンが閉じられていません"), isEscape)
    End Function

    ''' <summary>コードトークンリストを生成します。</summary>
    ''' <param name="reader">入力文字ストリーム。</param>
    ''' <returns>トークンリスト。</returns>
    Private Function CreateTokens(reader As StringPtr) As IToken
        Dim codeStr = GetPartitionString(reader, "コードトークンが閉じられていません")?.Trim()
        Dim lowStr = codeStr.ToLower()

        If lowStr.StartsWith("if ") Then
            Return New IfToken(SplitToken(codeStr.Substring(3)))
        ElseIf lowStr.StartsWith("else if ") Then
            Return New ElseIfToken(SplitToken(codeStr.Substring(8)))
        ElseIf lowStr.StartsWith("else") Then
            Return ElseToken.Value
        ElseIf lowStr.StartsWith("end if") Then
            Return EndIfToken.Value
        ElseIf lowStr.StartsWith("for each ") Then
            Return New ForEachToken(SplitToken(codeStr.Substring(9)))
        ElseIf lowStr.StartsWith("foreach ") Then
            Return New ForEachToken(SplitToken(codeStr.Substring(8)))
        ElseIf lowStr.StartsWith("end for") Then
            Return EndForToken.Value
        ElseIf lowStr.StartsWith("trim ") Then
            Return New TrimToken(codeStr.Substring(5))
        ElseIf lowStr.StartsWith("trim") Then
            Return New TrimToken()
        ElseIf lowStr.StartsWith("end trim") Then
            Return EndTrimToken.Value
        Else
            Return New QueryToken("")
        End If
        Return Nothing
    End Function

    ''' <summary>}で閉じられた文字列を取得します。</summary>
    ''' <param name="reader">入力文字ストリーム。</param>
    ''' <param name="errMessage">閉じられていないときのメッセージ。</param>
    ''' <returns>閉じられた範囲の文字列。</returns>
    Private Function GetPartitionString(reader As StringPtr, errMessage As String) As String
        Dim buffer As New StringBuilder()
        Do While reader.HasNext
            Dim c = reader.CurrentToMove()
            Select Case c
                Case "}"c
                    Return buffer.ToString()

                Case "\"c
                    If reader.Current = "{"c OrElse reader.Current = "}"c Then
                        buffer.Append(reader.Current)
                        reader.CurrentToMove()
                    End If

                Case Else
                    buffer.Append(c)
            End Select
        Loop
        Throw New DSqlAnalysisException(errMessage)
    End Function

    ''' <summary>文字列を分割してトークンリストを作成する。</summary>
    ''' <param name="input">対象文字列。</param>
    ''' <returns>トークンリスト。</returns>
    Function SplitToken(input As String) As List(Of TokenPosition)
        Dim keychar As New HashSet(Of Char)(New Char() {"+"c, "-"c, "*"c, "/"c, "("c, ")"c, "="c, "<"c, ">"c, "!"c, ChrW(0)})
        Dim tokens As New List(Of TokenPosition)()

        Dim reader = New StringPtr(input)

        Do While reader.HasNext
            Dim c = reader.Current()

            If Char.IsWhiteSpace(c) Then
                reader.Move(1)
            Else
                Dim pos = reader.CurrentPosition
                If reader.EqualKeyword("=") Then
                    tokens.Add(New TokenPosition(EqualToken.Value, pos))
                    reader.Move(1)
                ElseIf reader.EqualKeyword("<=") Then
                    tokens.Add(New TokenPosition(LessEqualToken.Value, pos))
                    reader.Move(2)
                ElseIf reader.EqualKeyword(">=") Then
                    tokens.Add(New TokenPosition(GreaterEqualToken.Value, pos))
                    reader.Move(2)
                ElseIf reader.EqualKeyword("<>") Then
                    tokens.Add(New TokenPosition(NotEqualToken.Value, pos))
                    reader.Move(2)
                ElseIf reader.EqualKeyword("!=") Then
                    tokens.Add(New TokenPosition(NotEqualToken.Value, pos))
                    reader.Move(2)
                ElseIf reader.EqualKeyword("<") Then
                    tokens.Add(New TokenPosition(LessToken.Value, pos))
                    reader.Move(1)
                ElseIf reader.EqualKeyword(">") Then
                    tokens.Add(New TokenPosition(GreaterToken.Value, pos))
                    reader.Move(1)
                ElseIf reader.EqualKeywordTail("true", keychar) Then
                    tokens.Add(New TokenPosition(TrueToken.Value, pos))
                    reader.Move(4)
                ElseIf reader.EqualKeywordTail("false", keychar) Then
                    tokens.Add(New TokenPosition(FalseToken.Value, pos))
                    reader.Move(5)
                ElseIf reader.EqualKeywordTail("not", keychar) Then
                    tokens.Add(New TokenPosition(NotToken.Value, pos))
                    reader.Move(3)
                ElseIf reader.EqualKeywordTail("null", keychar) Then
                    tokens.Add(New TokenPosition(NullToken.Value, pos))
                    reader.Move(4)
                ElseIf reader.EqualKeywordTail("nothing", keychar) Then
                    tokens.Add(New TokenPosition(NullToken.Value, pos))
                    reader.Move(7)
                ElseIf reader.EqualKeywordTail("and", keychar) Then
                    tokens.Add(New TokenPosition(AndToken.Value, pos))
                    reader.Move(3)
                ElseIf reader.EqualKeywordTail("or", keychar) Then
                    tokens.Add(New TokenPosition(OrToken.Value, pos))
                    reader.Move(2)
                ElseIf reader.EqualKeywordTail("in", keychar) Then
                    tokens.Add(New TokenPosition(InToken.Value, pos))
                    reader.Move(2)
                ElseIf reader.EqualKeyword("+") Then
                    tokens.Add(New TokenPosition(PlusToken.Value, pos))
                    reader.Move(1)
                ElseIf reader.EqualKeyword("-") Then
                    tokens.Add(New TokenPosition(MinusToken.Value, pos))
                    reader.Move(1)
                ElseIf reader.EqualKeyword("*") Then
                    tokens.Add(New TokenPosition(MultiToken.Value, pos))
                    reader.Move(1)
                ElseIf reader.EqualKeyword("/") Then
                    tokens.Add(New TokenPosition(DivToken.Value, pos))
                    reader.Move(1)
                ElseIf reader.EqualKeyword("(") Then
                    tokens.Add(New TokenPosition(LParenToken.Value, pos))
                    reader.Move(1)
                ElseIf reader.EqualKeyword(")") Then
                    tokens.Add(New TokenPosition(RParenToken.Value, pos))
                    reader.Move(1)
                ElseIf reader.EqualKeyword("!") Then
                    tokens.Add(New TokenPosition(NotToken.Value, pos))
                    reader.Move(1)
                ElseIf c = "#"c Then
                    tokens.Add(New TokenPosition(CreateDateToken(reader), pos))
                ElseIf c = "'"c Then
                    tokens.Add(New TokenPosition(CreateStringToken(reader), pos))
                ElseIf Char.IsDigit(c) Then
                    tokens.Add(New TokenPosition(CreateNumberToken(reader), pos))
                Else
                    Dim tkn = CreateIdentToken(reader, keychar)
                    If pos < reader.CurrentPosition Then
                        tokens.Add(New TokenPosition(tkn, pos))
                    Else
                        Throw New DSqlAnalysisException($"{pos}位置以降の文字列を判定できません")
                    End If
                End If
            End If
        Loop

        Return tokens
    End Function

    ''' <summary>日時トークンを生成します。</summary>
    ''' <param name="reader">入力文字ストリーム。</param>
    ''' <returns>日時トークン。</returns>
    Private Function CreateDateToken(reader As StringPtr) As ObjectToken
        Dim res As New StringBuilder()
        Dim closed = False

        reader.Move(1)

        Do While reader.HasNext
            Dim c = reader.Current()
            reader.Move(1)

            If c <> "#"c Then
                res.Append(c)
            Else
                closed = True
                Exit Do
            End If
        Loop

        If closed Then
            Dim lite = res.ToString()
            Dim dv As Date
            Dim tv As TimeSpan
            If TimeSpan.TryParse(lite, tv) Then
                Return New ObjectToken(tv)
            ElseIf Date.TryParse(lite, dv) Then
                Return New ObjectToken(dv)
            Else
                Throw New DSqlAnalysisException($"日付、時間が定義されていない")
            End If
        Else
            Throw New DSqlAnalysisException($"文字列リテラルが閉じられていない:{res}")
        End If
    End Function

    ''' <summary>文字列トークンを生成します。</summary>
    ''' <param name="reader">入力文字ストリーム。</param>
    ''' <returns>文字列トークン。</returns>
    Private Function CreateStringToken(reader As StringPtr) As StringToken
        Dim res As New StringBuilder()
        Dim closed = False

        reader.Move(1)

        Do While reader.HasNext
            Dim c = reader.Current()
            reader.Move(1)

            If c = "\"c AndAlso reader.Current = "'"c Then
                res.Append("'"c)
                reader.Move(1)
            ElseIf c = "'"c AndAlso reader.Current = "'"c Then
                res.Append("'"c)
                reader.Move(1)
            ElseIf c <> "'"c Then
                res.Append(c)
            Else
                closed = True
                Exit Do
            End If
        Loop

        If closed Then
            Return New StringToken(res.ToString())
        Else
            Throw New DSqlAnalysisException($"文字列リテラルが閉じられていない:{res}")
        End If
    End Function

    ''' <summary>数値トークンを生成します。</summary>
    ''' <param name="reader">入力文字ストリーム。</param>
    ''' <returns>数値トークン。</returns>
    Private Function CreateNumberToken(reader As StringPtr) As NumberToken
        Dim res As New StringBuilder()
        Dim dec = False

        Do While reader.HasNext
            Dim mc = reader.Current
            If Char.IsDigit(mc) Then
                reader.Move(1)
                res.Append(mc)
            ElseIf mc = "."c Then
                If Not dec Then
                    dec = True
                    reader.Move(1)
                    res.Append(mc)
                Else
                    res.Append(mc)
                    Throw New DSqlAnalysisException($"数値の変換ができません:{res}")
                End If
            Else
                Exit Do
            End If
        Loop

        Return NumberToken.Create(res.ToString())
    End Function

    ''' <summary>識別子トークンを生成します。</summary>
    ''' <param name="reader">入力文字ストリーム。</param>
    ''' <param name="keychar">識別子に使えない文字セット。</param>
    ''' <returns>識別子トークン。</returns>
    Private Function CreateIdentToken(reader As StringPtr, keychar As HashSet(Of Char)) As IToken
        Dim res As New StringBuilder()

        Do While reader.HasNext
            Dim c = reader.Current()

            If keychar.Contains(c) OrElse Char.IsWhiteSpace(c) Then
                Exit Do
            Else
                res.Append(c)
                reader.Move(1)
            End If
        Loop
        Return New IdentToken(res.ToString())
    End Function

    ''' <summary>入力文字ストリーム。</summary>
    Private NotInheritable Class StringPtr

        ''' <summary>シーク位置ポインタ。</summary>
        Private mPointer As Integer

        ''' <summary>入力文字配列。</summary>
        Private ReadOnly mChars As Char()

        ''' <summary>読み込みの終了していない文字があれば真を返す。</summary>
        Public ReadOnly Property HasNext As Boolean
            Get
                Return (Me.mPointer < If(Me.mChars?.Length, 0))
            End Get
        End Property

        ''' <summary>カレント文字を返す。</summary>
        Public ReadOnly Property Current As Char
            Get
                Return If(Me.mPointer < If(Me.mChars?.Length, 0), Me.mChars(Me.mPointer), ChrW(0))
            End Get
        End Property

        ''' <summary>カレント位置を返す。</summary>
        Public ReadOnly Property CurrentPosition As Integer
            Get
                Return Me.mPointer
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputstr">入力文字列。</param>
        Public Sub New(inputstr As String)
            Me.mPointer = 0
            Me.mChars = inputstr.ToCharArray()
        End Sub

        ''' <summary>カレント文字を取得し、ポイントをひとつ進める。</summary>
        ''' <returns>カレント文字列。</returns>
        Public Function CurrentToMove() As Char
            Dim c = Me.Current
            Me.mPointer += 1
            Return c
        End Function

        ''' <summary>カレント位置より相対的な文字を取得する。</summary>
        ''' <param name="point">相対位置。</param>
        ''' <returns>相対位置の文字。</returns>
        Public Function NestChar(point As Integer) As Char
            Dim p = Me.mPointer + point
            Return If(p < If(Me.mChars?.Length, 0), Me.mChars(p), ChrW(0))
        End Function

        ''' <summary>カレント位置から引数の文字列と一致したならば真を返す。</summary>
        ''' <param name="keyWord">判定する文字列。</param>
        ''' <returns>一致していたら真。</returns>
        Public Function EqualKeyword(keyWord As String) As Boolean
            Dim kcs = keyWord.ToCharArray()
            For i As Integer = 0 To kcs.Length - 1
                If kcs(i) <> Char.ToLower(Me.NestChar(i)) Then
                    Return False
                End If
            Next
            Return True
        End Function

        ''' <summary>カレント位置から引数の文字列と一致したならば真を返す（末尾チェックあり）</summary>
        ''' <param name="keyWord">判定する文字列。</param>
        ''' <param name="keychar">識別子に使えない文字セット。</param>
        ''' <returns>一致していたら真。</returns>
        Public Function EqualKeywordTail(keyWord As String, keychar As HashSet(Of Char)) As Boolean
            Dim kcs = keyWord.ToCharArray()
            For i As Integer = 0 To kcs.Length - 1
                If kcs(i) <> Char.ToLower(Me.NestChar(i)) Then
                    Return False
                End If
            Next
            Return keychar.Contains(Me.NestChar(kcs.Length)) OrElse Char.IsWhiteSpace(Me.NestChar(kcs.Length))
        End Function

        ''' <summary>カレント位置を移動させる。</summary>
        ''' <param name="moveAmount">移動量。</param>
        Public Sub Move(moveAmount As Integer)
            Me.mPointer += moveAmount
        End Sub

        '''' <summary>カレント位置から指定文字数の文字列を取得します。</summary>
        '''' <param name="len">文字数。</param>
        '''' <returns>文字列。</returns>
        'Public Function GetString(len As Integer) As String
        '    Dim buf As New StringBuilder()
        '    For i As Integer = Me.mPointer To Me.mPointer + len - 1
        '        buf.Append(Me.mChars(i))
        '    Next
        '    Return buf.ToString()
        'End Function

    End Class

End Module