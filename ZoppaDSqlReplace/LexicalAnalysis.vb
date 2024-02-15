Option Strict On
Option Explicit On

Imports System.Text
Imports ZoppaDSqlReplace.Tokens
Imports ZoppaDSqlReplace.TokenCollection
Imports System.Runtime.CompilerServices
Imports System.Net.Http

''' <summary>字句解析機能。</summary>
Friend Module LexicalAnalysis

    ''' <summary>トークン種別。</summary>
    Private Enum TKN_TYPE

        ''' <summary>クエリトークン。</summary>
        QUERY_TKN

        ''' <summary>コードトークン。</summary>
        CODE_TKN

        ''' <summary>置き換えトークン。</summary>
        REPLASE_TKN

        ''' <summary>置き換えトークン。</summary>
        REPLASE_TKN_ESC

    End Enum

    ''' <summary>文字列をトークン解析します。</summary>
    ''' <param name="command">文字列。</param>
    ''' <returns>トークンリスト。</returns>
    Public Function SplitQueryToken(command As String) As List(Of TokenPosition)
        Dim tokens As New List(Of TokenPosition)()

        Dim reader = New StringPtr(command)

        Dim buffer As New StringBuilder()
        Dim tknType = TKN_TYPE.QUERY_TKN
        Dim pos As Integer = 0
        Do While reader.HasNext
            Dim c = reader.Current()

            Select Case c
                Case "{"c
                    tokens.AddRange(buffer.CreateQueryTokens(pos))
                    pos = reader.CurrentPosition
                    reader.Move(1)
                    tknType = TKN_TYPE.CODE_TKN

                Case "}"c
                    Select Case tknType
                        Case TKN_TYPE.QUERY_TKN
                            Throw New DSqlAnalysisException("}に対応する{を入力してください")
                        Case TKN_TYPE.CODE_TKN
                            tokens.AddIfNull(buffer.CreateCodeToken(), pos)
                        Case TKN_TYPE.REPLASE_TKN
                            tokens.AddIfNull(buffer.CreateReplaseToken(False), pos)
                        Case TKN_TYPE.REPLASE_TKN_ESC
                            tokens.AddIfNull(buffer.CreateReplaseToken(True), pos)
                    End Select
                    pos = reader.CurrentPosition
                    reader.Move(1)
                    tknType = TKN_TYPE.QUERY_TKN

                Case "#"c, "!"c, "$"c
                    If reader.NestChar(1) = "{"c Then
                        tokens.AddRange(buffer.CreateQueryTokens(pos))
                        pos = reader.CurrentPosition
                        reader.Move(2)
                        tknType = If(c <> "#", TKN_TYPE.REPLASE_TKN, TKN_TYPE.REPLASE_TKN_ESC)
                    Else
                        buffer.Append(c)
                        reader.Move(1)
                    End If

                Case "\"c
                    If reader.NestChar(1) = "{"c OrElse reader.NestChar(1) = "}"c Then
                        reader.Move(1)
                        buffer.Append(reader.Current)
                        reader.Move(1)
                    Else
                        buffer.Append(c)
                        reader.Move(1)
                    End If

                Case Else
                    buffer.Append(c)
                    reader.Move(1)
            End Select
        Loop

        Select Case tknType
            Case TKN_TYPE.QUERY_TKN
                tokens.AddRange(buffer.CreateQueryTokens(pos))
            Case TKN_TYPE.CODE_TKN
                Throw New DSqlAnalysisException("コードトークンが閉じられていません")
            Case TKN_TYPE.REPLASE_TKN, TKN_TYPE.REPLASE_TKN_ESC
                Throw New DSqlAnalysisException("置き換えトークンが閉じられていません")
        End Select

        Return tokens
    End Function

    ''' <summary>改行単位で不要な空白トークンを削除します。</summary>
    ''' <param name="srcTokens">トークンリスト。</param>
    ''' <returns>空白を取り除いたトークンリスト。</returns>
    Public Function RemoveSpaceCrLf(srcTokens As List(Of TokenPosition)) As List(Of TokenPosition)
        Dim res As New List(Of TokenPosition)

        Dim buffer As New List(Of TokenPosition)

        ' 改行単位でトークンを収集して調整する
        Dim pointer As New TokenStream(srcTokens)
        Do While pointer.HasNext
            buffer.Add(pointer.Current)

            ' 改行を見つけたので調整する
            If pointer.Current.IsCrLf Then
                AjustLineTokens(res, buffer)
                buffer.Clear()
            End If

            pointer.Move(1)
        Loop

        AjustLineTokens(res, buffer)
        Return res
    End Function

    ''' <summary>改行単位の空白削除を実施します。</summary>
    ''' <param name="res">戻り値のトークンリスト。</param>
    ''' <param name="buffer">調整対象のトークンリスト</param>
    Private Sub AjustLineTokens(res As List(Of TokenPosition), buffer As List(Of TokenPosition))
        ' 直接トークン、置き換えトークンを含む行か判定
        Dim rep = False
        For i As Integer = 0 To buffer.Count - 1
            If Not buffer(i).IsWhiteSpace AndAlso Not buffer(i).IsCrLf Then
                If buffer(i).TokenType = GetType(QueryToken) OrElse
                   buffer(i).TokenType = GetType(ReplaseToken) Then
                    rep = True
                    Exit For
                End If
            End If
        Next

        If rep Then
            ' 直接トークン、置き換えトークンを含む行はそのまま追加
            res.AddRange(buffer)
        Else
            ' 直接トークン、置き換えトークンを含まない行は前方の空白トークンを削除して追加
            '
            ' 1. 直前の行の末尾の改行を削除
            ' 2. 前方の空白を削除
            For i As Integer = 0 To buffer.Count - 1
                If Not buffer(i).IsWhiteSpace Then
                    If res.Count > 0 Then res.RemoveAt(res.Count - 1)   ' 1

                    For j As Integer = i To buffer.Count - 1            ' 2
                        res.Add(buffer(j))
                    Next
                    Exit For
                End If
            Next
        End If
    End Sub

    ''' <summary>トークンリストにトークンを追加します(Nullを除く)</summary>
    ''' <param name="tokens">トークンリスト。</param>
    ''' <param name="token">追加するトークン。</param>
    ''' <param name="pos">トークン位置。</param>
    <Extension>
    Private Sub AddIfNull(tokens As List(Of TokenPosition), token As IToken, pos As Integer)
        If token IsNot Nothing Then
            tokens.Add(New TokenPosition(token, pos))
        End If
    End Sub

    ''' <summary>直接出力するクエリトークンを生成します。</summary>
    ''' <param name="buffer">文字列バッファ。</param>
    ''' <param name="startPos">トークンの最初の文字位置。</param>
    ''' <returns>直接出力するクエリトークンを生成します。</returns>
    <Extension>
    Private Function CreateQueryTokens(buffer As StringBuilder, startPos As Integer) As List(Of TokenPosition)
        Dim res As New List(Of TokenPosition)
        If buffer.Length > 0 Then
            ' 一文字目の文字種を取得
            Dim preKind = QueryToken.GetCharKind(buffer(0))

            ' 一文字目をバッファに追加
            Dim buf As New StringBuilder()
            buf.Append(buffer(0))

            ' 文字種の変わり目でトークンを分割
            For i As Integer = 1 To buffer.Length - 1
                Dim kind = QueryToken.GetCharKind(buffer(i))

                ' 文字種が変わったのでトークンを分割
                If kind <> preKind Then
                    res.Add(New TokenPosition(New QueryToken(buf.ToString(), preKind), startPos))
                    buf.Clear()
                    startPos = i
                    preKind = kind
                End If

                buf.Append(buffer(i))
            Next
            res.Add(New TokenPosition(New QueryToken(buf.ToString(), preKind), startPos))
        End If

        buffer.Clear()
        Return res
    End Function

    ''' <summary>コードトークンリストを生成します。</summary>
    ''' <param name="buffer">文字列バッファ。</param>
    ''' <returns>コードトークン</returns>
    <Extension>
    Private Function CreateCodeToken(buffer As StringBuilder) As IToken
        Dim codeStr = buffer.ToString().Trim()
        buffer.Clear()
        Dim lowStr = If(codeStr.Length > 10, codeStr.Substring(0, 10), codeStr).ToLower()

        If lowStr.StartsWith("if ") Then
            Return New IfToken(SplitToken(codeStr.Substring(3)))
        ElseIf lowStr.StartsWith("else if ") Then
            Return New ElseIfToken(SplitToken(codeStr.Substring(8)))
        ElseIf lowStr.StartsWith("else") Then
            Return ElseToken.Value
        ElseIf lowStr.StartsWith("end if") Then
            Return EndIfToken.Value
        ElseIf lowStr.StartsWith("/if") Then
            Return EndIfToken.Value
        ElseIf lowStr.StartsWith("for each ") Then
            Return New ForEachToken(SplitToken(codeStr.Substring(9)))
        ElseIf lowStr.StartsWith("foreach ") Then
            Return New ForEachToken(SplitToken(codeStr.Substring(8)))
        ElseIf lowStr.StartsWith("end for") Then
            Return EndForToken.Value
        ElseIf lowStr.StartsWith("/for") Then
            Return EndForToken.Value
        ElseIf lowStr.StartsWith("trim where ") Then
            Return New TrimToken(False, True, codeStr.Substring(10))
        ElseIf lowStr.StartsWith("trim where") Then
            Return New TrimToken(False, True)
        ElseIf lowStr.StartsWith("trim both ") Then
            Return New TrimToken(True, False, codeStr.Substring(9))
        ElseIf lowStr.StartsWith("trim both") Then
            Return New TrimToken(True, False)
        ElseIf lowStr.StartsWith("trim ") Then
            Return New TrimToken(False, False, codeStr.Substring(5))
        ElseIf lowStr.StartsWith("trim") Then
            Return New TrimToken(False, False)
        ElseIf lowStr.StartsWith("end trim") Then
            Return EndTrimToken.Value
        ElseIf lowStr.StartsWith("/trim") Then
            Return EndTrimToken.Value
        ElseIf lowStr.StartsWith("select ") Then
            Return New SelectToken(SplitToken(codeStr.Substring(7)))
        ElseIf lowStr.StartsWith("case ") Then
            Return New CaseToken(SplitToken(codeStr.Substring(5)))
        ElseIf lowStr.StartsWith("end select") Then
            Return EndSelectToken.Value
        ElseIf lowStr.StartsWith("/select") Then
            Return EndSelectToken.Value
        End If
        Return Nothing
    End Function

    ''' <summary>置き換えトークンを生成します。</summary>
    ''' <param name="buffer">文字列バッファ。</param>
    ''' <param name="isEscape">エスケープするならば真。</param>
    ''' <returns>置き換えトークン。</returns>
    <Extension>
    Private Function CreateReplaseToken(buffer As StringBuilder, isEscape As Boolean) As IToken
        Dim repStr = buffer.ToString().Trim()
        buffer.Clear()
        Return If(repStr.Length > 0, New ReplaseToken(repStr, isEscape), Nothing)
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
                ElseIf reader.EqualKeyword(",") Then
                    tokens.Add(New TokenPosition(CommaToken.Value, pos))
                    reader.Move(1)
                ElseIf c = "#"c Then
                    tokens.Add(New TokenPosition(CreateDateToken(reader), pos))
                ElseIf c = "'"c Then
                    tokens.Add(New TokenPosition(CreateStringToken(reader, "'"c), pos))
                ElseIf c = """"c Then
                    tokens.Add(New TokenPosition(CreateStringToken(reader, """"c), pos))
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
    ''' <param name="bktChar">囲み文字。</param>
    ''' <returns>文字列トークン。</returns>
    Private Function CreateStringToken(reader As StringPtr, bktChar As Char) As StringToken
        Dim res As New StringBuilder()
        Dim closed = False

        reader.Move(1)

        Do While reader.HasNext
            Dim c = reader.Current()
            reader.Move(1)

            If c = "\"c AndAlso reader.Current = bktChar Then
                res.Append(bktChar)
                reader.Move(1)
            ElseIf c = bktChar AndAlso reader.Current = bktChar Then
                res.Append(bktChar)
                reader.Move(1)
            ElseIf c <> bktChar Then
                res.Append(c)
            Else
                closed = True
                Exit Do
            End If
        Loop

        If closed Then
            Return New StringToken(res.ToString(), bktChar)
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

    End Class

End Module