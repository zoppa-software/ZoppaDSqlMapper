Option Strict On
Option Explicit On

Imports System.Text
Imports ZoppaDSqlReplace.Environments
Imports ZoppaDSqlReplace.Express
Imports ZoppaDSqlReplace.Tokens
Imports ZoppaDSqlReplace.TokenCollection
Imports System.Security.Cryptography

''' <summary>トークンを解析する。</summary>
Friend Module ParserAnalysis

    ''' <summary>SQLの置き換えを実施します。</summary>
    ''' <param name="sqlQuery">元のSQL。</param>
    ''' <param name="tokens">トークンリスト。</param>
    ''' <param name="parameter">環境値情報。</param>
    ''' <returns>置き換え結果。</returns>
    Friend Function Replase(sqlQuery As String,
                            tokens As List(Of TokenPosition),
                            parameter As IEnvironmentValue) As String
        ' SQLの置き換えを実施
        Dim buffer As New StringBuilder()
        ReplaseQuery(sqlQuery, New TokenStream(tokens), parameter, buffer)

        ' 空白行を削除
        Return RemoveSpaceLine(buffer).ToString()
    End Function

#Region "ajust token list"

    ''' <summary>空白行を削除します。</summary>
    ''' <param name="buffer">対象バッファ。</param>
    ''' <returns>空白行を削除したバッファ。</returns>
    Private Function RemoveSpaceLine(buffer As StringBuilder) As StringBuilder
        Dim res As New StringBuilder(buffer.Length)

        ' 改行単位で分割
        Dim lines = buffer.ToString().Split(New String() {Environment.NewLine}, StringSplitOptions.None)

        ' 空白行以外を追加
        For i As Integer = 0 To lines.Length - 2
            If lines(i).Trim() <> "" Then
                res.AppendLine(lines(i))
            End If
        Next

        ' 最終行を追加
        If lines(lines.Length - 1).Trim() <> "" Then
            res.Append(lines(lines.Length - 1))
        End If
        Return res
    End Function

#End Region

    ''' <summary>SQLの置き換えを実施します。</summary>
    ''' <param name="sqlQuery">元のSQL。</param>
    ''' <param name="reader">トークンリーダー。</param>
    ''' <param name="parameter">環境値情報。</param>
    Private Sub ReplaseQuery(sqlQuery As String,
                             reader As TokenStream,
                             parameter As IEnvironmentValue,
                             buffer As StringBuilder)
        Do While reader.HasNext
            Dim tkn = reader.Current

            Select Case tkn.TokenType
                Case GetType(QueryToken)
                    buffer.Append(tkn.ToString())
                    reader.Move(1)

                Case GetType(ReplaseToken)
                    Dim rtoken = tkn.GetToken(Of ReplaseToken)()
                    If rtoken IsNot Nothing Then
                        Dim rval = parameter.GetValue(If(rtoken.Contents?.ToString(), ""))
                        Dim ans = GetRefValue(rval, rtoken.IsEscape)
                        buffer.Append(ans)
                    End If
                    reader.Move(1)

                Case GetType(IfToken)
                    reader.Move(1)
                    Dim ifTokens = CollectBlockToken(sqlQuery, reader, GetType(IfToken), GetType(EndIfToken))
                    EvaluationIf(sqlQuery, tkn.GetToken(Of IfToken)(), ifTokens, buffer, parameter)

                Case GetType(ElseIfToken), GetType(ElseToken), GetType(EndIfToken)
                    Throw New DSqlAnalysisException($"ifが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")

                Case GetType(ForEachToken)
                    reader.Move(1)
                    Dim forTokens = CollectBlockToken(sqlQuery, reader, GetType(ForEachToken), GetType(EndForToken))
                    EvaluationFor(sqlQuery, tkn.GetToken(Of ForEachToken)(), forTokens, buffer, parameter)

                Case GetType(EndForToken)
                    Throw New DSqlAnalysisException($"forが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")

                Case GetType(SelectToken)
                    reader.Move(1)
                    Dim selTokens = CollectBlockToken(sqlQuery, reader, GetType(SelectToken), GetType(EndSelectToken))
                    EvaluationSelect(sqlQuery, tkn.GetToken(Of SelectToken)(), selTokens, buffer, parameter)

                Case GetType(CaseToken), GetType(EndSelectToken)
                    Throw New DSqlAnalysisException($"selectが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")

                Case GetType(TrimToken)
                    reader.Move(1)
                    Dim trimTokens = CollectBlockToken(sqlQuery, reader, GetType(TrimToken), GetType(EndTrimToken))
                    EvaluationTrim(sqlQuery, tkn.GetToken(Of TrimToken)(), trimTokens, buffer, parameter)

                Case GetType(EndTrimToken)
                    Throw New DSqlAnalysisException($"trimが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")
            End Select
        Loop
    End Sub

#Region "token manage"

    ''' <summary>ブロックトークンを集めます。</summary>
    ''' <param name="sqlQuery">元のSQL。</param>
    ''' <param name="reader">トークンリーダー。</param>
    ''' <param name="startTokenName">開始トークン。</param>
    ''' <param name="endTokenName">終了トークン。</param>
    ''' <returns>ブロックトークン。</returns>
    Private Function CollectBlockToken(sqlQuery As String,
                                       reader As TokenStream,
                                       startTokenName As Type,
                                       endTokenName As Type) As List(Of TokenPosition)
        Dim res As New List(Of TokenPosition)()
        Dim startToken = reader.Current

        Dim nest As Integer = 0
        Do While reader.HasNext
            Dim tkn = reader.Current
            If tkn.TokenType Is startTokenName Then
                nest += 1
            ElseIf tkn.TokenType Is endTokenName Then
                If nest = 0 Then
                    reader.Move(1)
                    Return res
                Else
                    nest -= 1
                End If
            End If
            res.Add(tkn)
            reader.Move(1)
        Loop

        Throw New DSqlAnalysisException($"フロックが閉じられていません。{vbCrLf}{sqlQuery}:{startToken.Position}")
    End Function

    ''' <summary>Ifを評価します。</summary>
    ''' <param name="sqlQuery">元のSQL。</param>
    ''' <param name="sifToken">Ifのトークンリスト。</param>
    ''' <param name="tokens">ブロック内のトークンリスト。</param>
    ''' <param name="buffer">結果バッファ。</param>
    ''' <param name="parameter">パラメータ。</param>
    Private Sub EvaluationIf(sqlQuery As String,
                             sifToken As IToken,
                             tokens As List(Of TokenPosition),
                             buffer As StringBuilder,
                             parameter As IEnvironmentValue)
        Dim blocks As New List(Of (condition As IToken, block As List(Of TokenPosition))) From {
            (sifToken, New List(Of TokenPosition)())
        }

        ' If、ElseIf、Elseブロックを集める
        Dim nest As Integer = 0
        For Each tkn In tokens
            Select Case tkn.TokenType
                Case GetType(IfToken)
                    nest += 1
                    blocks(blocks.Count - 1).block.Add(tkn)

                Case GetType(ElseIfToken)
                    If nest = 0 Then
                        blocks.Add((tkn.GetToken(Of ElseIfToken), New List(Of TokenPosition)()))
                    Else
                        blocks(blocks.Count - 1).block.Add(tkn)
                    End If

                Case GetType(ElseToken)
                    If nest = 0 Then
                        blocks.Add((tkn.GetToken(Of ElseToken), New List(Of TokenPosition)()))
                    Else
                        blocks(blocks.Count - 1).block.Add(tkn)
                    End If

                Case GetType(EndIfToken)
                    If nest = 0 Then
                        Exit For
                    Else
                        nest -= 1
                        blocks(blocks.Count - 1).block.Add(tkn)
                    End If

                Case Else
                    blocks(blocks.Count - 1).block.Add(tkn)
            End Select
        Next

        ' If、ElseIf、Elseブロックを評価
        Dim lclbuf As New StringBuilder()
        For Each tkn In blocks
            Select Case tkn.condition.TokenType
                Case GetType(IfToken), GetType(ElseIfToken)
                    ' 条件を評価して真ならば、ブロックを出力
                    Dim ifans = Executes(DirectCast(tkn.condition, ICommandToken).CommandTokens, parameter)
                    If TypeOf ifans.Contents Is Boolean AndAlso CBool(ifans.Contents) Then
                        Dim tkns As New List(Of TokenPosition)(tkn.block)
                        ReplaseQuery(sqlQuery, New TokenStream(tkns), parameter, lclbuf)
                        Exit For
                    End If

                Case GetType(ElseToken)
                    Dim tkns As New List(Of TokenPosition)(tkn.block)
                    ReplaseQuery(sqlQuery, New TokenStream(tkns), parameter, lclbuf)
            End Select
        Next
        buffer.Append(lclbuf.ToString())
    End Sub

    ''' <summary>Foreachを評価します。</summary>
    ''' <param name="sqlQuery">元のSQL。</param>
    ''' <param name="sforToken">Forのトークンリスト。</param>
    ''' <param name="tokens">ブロック内のトークンリスト。</param>
    ''' <param name="buffer">結果バッファ。</param>
    ''' <param name="parameter">パラメータ。</param>
    Private Sub EvaluationFor(sqlQuery As String,
                              sforToken As ForEachToken,
                              tokens As List(Of TokenPosition),
                              buffer As StringBuilder,
                              parameter As IEnvironmentValue)
        ' カウンタ変数
        Dim valkey As String = ""

        ' ループ元コレクション
        Dim collection As IEnumerable = Nothing

        ' 構文を解析して変数、ループ元コレクションを取得
        With sforToken
            If .CommandTokens.Count = 3 AndAlso
               .CommandTokens(0).TokenType Is GetType(IdentToken) AndAlso
               .CommandTokens(1).TokenType Is GetType(InToken) AndAlso
               .CommandTokens(2).TokenType Is GetType(IdentToken) Then
                Dim colln = parameter.GetValue(If(.CommandTokens(2).Contents?.ToString(), ""))
                If TypeOf colln Is IEnumerable Then
                    valkey = If(.CommandTokens(0).Contents?.ToString(), "")
                    collection = CType(colln, IEnumerable)
                End If
            End If
        End With

        ' Foreachして出力
        Dim otherParameter = parameter.Clone()
        For Each v In collection
            otherParameter.AddVariant(valkey, v)

            Dim lclbuf As New StringBuilder()
            ReplaseQuery(sqlQuery, New TokenStream(tokens), otherParameter, lclbuf)
            buffer.Append(lclbuf.ToString())
        Next
    End Sub

    ''' <summary>Trimを評価します。</summary>
    ''' <param name="sqlQuery">元のSQL。</param>
    ''' <param name="strimToken">Trimのトークンリスト。</param>
    ''' <param name="tokens">ブロック内のトークンリスト。</param>
    ''' <param name="buffer">結果バッファ。</param>
    ''' <param name="parameter">パラメータ。</param>
    Private Sub EvaluationTrim(sqlQuery As String,
                               strimToken As TrimToken,
                               tokens As List(Of TokenPosition),
                               buffer As StringBuilder,
                               parameter As IEnvironmentValue)
        Dim lclbuf As New StringBuilder()
        ReplaseQuery(sqlQuery, New TokenStream(tokens), parameter, lclbuf)

        Dim tarStr = lclbuf.ToString()

        For Each trimWord In strimToken.TrimStrings
            If tarStr.TrimEnd().EndsWith(trimWord, StringComparison.CurrentCultureIgnoreCase) Then
                Dim idx = tarStr.LastIndexOf(trimWord, StringComparison.CurrentCultureIgnoreCase)
                tarStr = tarStr.Remove(idx, trimWord.Length).TrimEnd(" "c, vbTab(0))
                Exit For
            End If
        Next

        If strimToken.IsBoth Then
            For Each trimWord In strimToken.TrimStrings
                If tarStr.TrimStart().StartsWith(trimWord, StringComparison.CurrentCultureIgnoreCase) Then
                    Dim idx = tarStr.IndexOf(trimWord, StringComparison.CurrentCultureIgnoreCase)
                    tarStr = tarStr.Remove(idx, trimWord.Length).TrimStart(" "c, vbTab(0))
                    Exit For
                End If
            Next
        End If

        buffer.Append(tarStr)
    End Sub

    ''' <summary>Selectを評価します。</summary>
    ''' <param name="sqlQuery">元のSQL。</param>
    ''' <param name="sselToken">Ifのトークンリスト。</param>
    ''' <param name="tokens">ブロック内のトークンリスト。</param>
    ''' <param name="buffer">結果バッファ。</param>
    ''' <param name="parameter">パラメータ。</param>
    Private Sub EvaluationSelect(sqlQuery As String,
                                 sselToken As IToken,
                                 tokens As List(Of TokenPosition),
                                 buffer As StringBuilder,
                                 parameter As IEnvironmentValue)
        Dim blocks As New List(Of (condition As IToken, block As List(Of TokenPosition))) From {
            (sselToken, New List(Of TokenPosition)())
        }

        ' Select、Case、Elseブロックを集める
        Dim nest As Integer = 0
        For Each tkn In tokens
            Select Case tkn.TokenType
                Case GetType(SelectToken)
                    nest += 1
                    blocks(blocks.Count - 1).block.Add(tkn)

                Case GetType(CaseToken)
                    If nest = 0 Then
                        blocks.Add((tkn.GetToken(Of CaseToken), New List(Of TokenPosition)()))
                    Else
                        blocks(blocks.Count - 1).block.Add(tkn)
                    End If

                Case GetType(ElseToken)
                    If nest = 0 Then
                        blocks.Add((tkn.GetToken(Of ElseToken), New List(Of TokenPosition)()))
                    Else
                        blocks(blocks.Count - 1).block.Add(tkn)
                    End If

                Case GetType(EndSelectToken)
                    If nest = 0 Then
                        Exit For
                    Else
                        nest -= 1
                        blocks(blocks.Count - 1).block.Add(tkn)
                    End If

                Case Else
                    blocks(blocks.Count - 1).block.Add(tkn)
            End Select
        Next

        ' Select、Case、Elseブロックを評価
        Dim lclbuf As New StringBuilder()

        Dim caseVal = Executes(DirectCast(blocks(0).condition, ICommandToken).CommandTokens, parameter)
        For i As Integer = 1 To blocks.Count - 1
            Dim tkn = blocks(i)

            Select Case tkn.condition.TokenType
                Case GetType(CaseToken)
                    ' 条件を評価して真ならば、ブロックを出力
                    Dim caseans = Executes(DirectCast(tkn.condition, ICommandToken).CommandTokens, parameter)
                    If caseans.Contents?.Equals(caseVal.Contents) Then
                        Dim tkns As New List(Of TokenPosition)(tkn.block)
                        ReplaseQuery(sqlQuery, New TokenStream(tkns), parameter, lclbuf)
                        Exit For
                    End If

                Case GetType(ElseToken)
                    Dim tkns As New List(Of TokenPosition)(tkn.block)
                    ReplaseQuery(sqlQuery, New TokenStream(tkns), parameter, lclbuf)
            End Select
        Next
        'For Each tkn In blocks
        '    Select Case tkn.condition.TokenType
        '        Case GetType(IfToken), GetType(ElseIfToken)
        '            ' 条件を評価して真ならば、ブロックを出力
        '            Dim ifans = Executes(DirectCast(tkn.condition, ICommandToken).CommandTokens, parameter)
        '            If TypeOf ifans.Contents Is Boolean AndAlso CBool(ifans.Contents) Then
        '                Dim tkns As New List(Of TokenPosition)(tkn.block)
        '                ReplaseQuery(sqlQuery, New TokenStream(tkns), parameter, lclbuf)
        '                Exit For
        '            End If

        '        Case GetType(ElseToken)
        '            Dim tkns As New List(Of TokenPosition)(tkn.block)
        '            ReplaseQuery(sqlQuery, New TokenStream(tkns), parameter, lclbuf)
        '    End Select
        'Next
        buffer.Append(lclbuf.ToString())
    End Sub

    ''' <summary>パラメータ値を参照して取得します。</summary>
    ''' <param name="refObj">参照オブジェクト。</param>
    ''' <param name="isEscape">エスケープしているならば真。</param>
    ''' <returns>取得した値を表現する文字列。</returns>
    Private Function GetRefValue(refObj As Object, isEscape As Boolean) As String
        If TypeOf refObj Is String AndAlso isEscape Then
            ' エスケープして文字列を取得
            Dim s = refObj.ToString()
            s = s.Replace("'"c, "''")
            s = s.Replace("\"c, "\\")
            Return $"'{s}'"

        ElseIf TypeOf refObj Is String Then
            ' 文字列を取得
            Return refObj.ToString()

        ElseIf TypeOf refObj Is IEnumerable Then
            ' 列挙して値を取得
            Dim buf As New StringBuilder()
            For Each itm In CType(refObj, IEnumerable)
                If buf.Length > 0 Then
                    buf.Append(", ")
                End If
                buf.Append(GetRefValue(itm, isEscape))
            Next
            Return buf.ToString()

        ElseIf refObj Is Nothing Then
            ' null値を取得
            Return "null"
        Else
            Return refObj.ToString()
        End If
    End Function

#End Region

    ''' <summary>式を解析して結果を取得します。</summary>
    ''' <param name="tokens">対象トークン。。</param>
    ''' <param name="parameter">パラメータ。</param>
    ''' <returns>解析結果。</returns>
    Friend Function Executes(tokens As List(Of TokenPosition), parameter As IEnvironmentValue) As IToken
        ' 式木を作成
        Dim logicalParser As New LogicalParser()
        Dim compParser As New ComparisonParser()
        Dim addOrSubParser As New AddOrSubParser()
        Dim multiOrDivParser As New MultiOrDivParser()
        Dim facParser As New FactorParser()
        Dim parenParser As New ParenParser()

        ' 解析クラスを構成
        logicalParser.NextParser = compParser
        compParser.NextParser = addOrSubParser
        addOrSubParser.NextParser = multiOrDivParser
        multiOrDivParser.NextParser = facParser
        facParser.NextParser = parenParser
        parenParser.NextParser = logicalParser

        ' トークン解析
        Dim tknPtr = New TokenStream(tokens)
        Dim expr = logicalParser.Parser(tknPtr)

        ' 結果を取得する
        If Not tknPtr.HasNext Then
            Return expr.Executes(parameter)
        Else
            Throw New DSqlAnalysisException("未評価のトークンがあります")
        End If
    End Function

    ''' <summary>括弧内部式を取得します。</summary>
    ''' <param name="reader">入力トークンストリーム。</param>
    ''' <param name="nxtParser">次のパーサー。</param>
    ''' <returns>括弧内部式。</returns>
    Private Function CreateParenExpress(reader As TokenStream, nxtParser As IParser) As ParenExpress
        Dim tmp As New List(Of TokenPosition)()
        Dim lv As Integer = 0
        Do While reader.HasNext
            Dim tkn = reader.Current
            reader.Move(1)

            Select Case tkn.TokenType
                Case GetType(LParenToken)
                    tmp.Add(tkn)
                    lv += 1

                Case GetType(RParenToken)
                    If lv > 0 Then
                        tmp.Add(tkn)
                        lv -= 1
                    Else
                        Exit Do
                    End If

                Case Else
                    tmp.Add(tkn)
            End Select
        Loop
        Return New ParenExpress(nxtParser.Parser(New TokenStream(tmp)))
    End Function

#Region "express classes"

    ''' <summary>評価部分ポインタです。</summary>
    Private NotInheritable Class EvaPartsPointer
        Implements IEnumerable(Of EvaParts)

        ' 元のリスト
        Private ReadOnly mParts As List(Of EvaParts)

        ' インデックス
        Private mIndex As Integer = 0

        ''' <summary>指定位置の評価要素を取得。</summary>
        ''' <param name="idx">インデックス。</param>
        Default Public ReadOnly Property Items(idx As Integer) As EvaParts
            Get
                Return Me.mParts(idx)
            End Get
        End Property

        ''' <summary>現在のカレントのインデックスを取得。</summary>
        Public ReadOnly Property Index As Integer
            Get
                Return Me.mIndex
            End Get
        End Property

        ''' <summary>要素数を取得。</summary>
        Public ReadOnly Property Count As Integer
            Get
                Return If(Me.mParts?.Count, 0)
            End Get
        End Property

        ''' <summary>カレントの評価部分を取得します。</summary>
        ''' <returns>評価部分。</returns>
        Public ReadOnly Property Current As EvaParts
            Get
                Return Me.mParts(Me.mIndex)
            End Get
        End Property

        ''' <summary>ポイントに残りがあれば真を返す。</summary>
        Public ReadOnly Property HasNext As Boolean
            Get
                Return Me.mIndex < Me.mParts.Count
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="srcParts">ポイントするリスト。</param>
        Public Sub New(srcParts As List(Of EvaParts))
            Me.mParts = New List(Of EvaParts)(srcParts)
        End Sub

        ''' <summary>ポイントをインクリメントします。</summary>
        Public Sub Increment()
            Me.mIndex += 1
        End Sub

        ''' <summary>先頭が指定の文字列ならば指定文字列をスキップします。</summary>
        ''' <param name="word">スキップする文字列。</param>
        Public Sub SkipWord(word As String)
            Do While Me.HasNext
                If Me.Current.IsSpace Then
                    Me.Increment()
                ElseIf Me.Current.OutString.Trim().ToLower().StartsWith(word) Then
                    Me.Increment()
                    Exit Do
                Else
                    Exit Do
                End If
            Loop
        End Sub

        ''' <summary>列挙子を取得します。</summary>
        ''' <returns>列挙子。</returns>
        Public Function GetEnumerator() As IEnumerator(Of EvaParts) Implements IEnumerable(Of EvaParts).GetEnumerator
            Return Me.mParts.GetEnumerator()
        End Function

        ''' <summary>列挙子を取得します。</summary>
        ''' <returns>列挙子。</returns>
        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return Me.GetEnumerator()
        End Function

    End Class

    ''' <summary>評価部分です。</summary>
    Private NotInheritable Class EvaParts

        ''' <summary>トークン位置を取得します。</summary>
        ''' <returns>トークン位置。</returns>
        Public ReadOnly Property TokenPos As TokenPosition

        ''' <summary>出力するならば真。</summary>
        Public Property IsOutpit As Boolean

        ''' <summary>出力する文字列です。</summary>
        Public Property OutString As String

        ''' <summary>コントロールトークンならば真を返します。</summary>
        ''' <returns>コントロールトークンならば真。</returns>
        Public ReadOnly Property IsControlToken As Boolean

        ''' <summary>空白文字列を出力するならば真を返します。</summary>
        Public ReadOnly Property IsSpace As Boolean
            Get
                Return (Me.OutString?.Trim() = "")
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tkn">トークン位置。</param>
        Public Sub New(tkn As TokenPosition)
            Me.TokenPos = tkn
            Me.IsOutpit = False
            Me.OutString = ""
            Me.IsControlToken = Me.TokenPos.IsControlToken
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tkn">トークン位置。</param>
        ''' <param name="outStr">出力文字列。</param>
        Public Sub New(tkn As TokenPosition, outStr As String)
            Me.TokenPos = tkn
            Me.IsOutpit = True
            Me.OutString = outStr
            Me.IsControlToken = Me.TokenPos.IsControlToken
        End Sub

        ''' <summary>文字列表現を取得します。</summary>
        Public Overrides Function ToString() As String
            Return $"{Me.OutString} view:{Me.IsOutpit} ctrl:{Me.IsControlToken}"
        End Function

    End Class

    ''' <summary>解析インターフェイス。</summary>
    Private Interface IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Function Parser(reader As TokenStream) As IExpression

    End Interface

    ''' <summary>括弧解析。</summary>
    Private NotInheritable Class ParenParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tkn = reader.Current
            If tkn.TokenType Is GetType(LParenToken) Then
                reader.Move(1)
                Return CreateParenExpress(reader, Me.NextParser)
            Else
                Return Me.NextParser.Parser(reader)
            End If
        End Function

    End Class

    ''' <summary>論理解析。</summary>
    Private NotInheritable Class LogicalParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            Do While reader.HasNext
                Dim ope = reader.Current
                Select Case ope.TokenType
                    Case GetType(AndToken)
                        reader.Move(1)
                        tml = New AndExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(OrToken)
                        reader.Move(1)
                        tml = New OrExpress(tml, Me.NextParser.Parser(reader))

                    Case Else
                        Exit Do
                End Select
            Loop

            Return tml
        End Function

    End Class

    ''' <summary>比較解析。</summary>
    Private NotInheritable Class ComparisonParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            If reader.HasNext Then
                Dim ope = reader.Current
                Select Case ope.TokenType
                    Case GetType(EqualToken)
                        reader.Move(1)
                        tml = New EqualExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(NotEqualToken)
                        reader.Move(1)
                        tml = New NotEqualExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(GreaterToken)
                        reader.Move(1)
                        tml = New GreaterExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(GreaterEqualToken)
                        reader.Move(1)
                        tml = New GreaterEqualExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(LessToken)
                        reader.Move(1)
                        tml = New LessExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(LessEqualToken)
                        reader.Move(1)
                        tml = New LessEqualExpress(tml, Me.NextParser.Parser(reader))
                End Select
            End If

            Return tml
        End Function

    End Class

    ''' <summary>加算、減算解析。</summary>
    Private NotInheritable Class AddOrSubParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            Do While reader.HasNext
                Dim ope = reader.Current
                Select Case ope.TokenType
                    Case GetType(PlusToken)
                        reader.Move(1)
                        tml = New PlusExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(MinusToken)
                        reader.Move(1)
                        tml = New MinusExpress(tml, Me.NextParser.Parser(reader))

                    Case Else
                        Exit Do
                End Select
            Loop

            Return tml
        End Function

    End Class

    ''' <summary>乗算、除算解析。</summary>
    Private NotInheritable Class MultiOrDivParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            Do While reader.HasNext
                Dim ope = reader.Current
                Select Case ope.TokenType
                    Case GetType(MultiToken)
                        reader.Move(1)
                        tml = New MultiExpress(tml, Me.NextParser.Parser(reader))

                    Case GetType(DivToken)
                        reader.Move(1)
                        tml = New DivExpress(tml, Me.NextParser.Parser(reader))

                    Case Else
                        Exit Do
                End Select
            Loop

            Return tml
        End Function
    End Class

    ''' <summary>要素解析。</summary>
    Private NotInheritable Class FactorParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tkn = reader.Current

            Select Case tkn.TokenType
                Case GetType(IdentToken), GetType(NumberToken), GetType(StringToken),
                     GetType(QueryToken), GetType(ReplaseToken), GetType(ObjectToken),
                     GetType(TrueToken), GetType(FalseToken), GetType(NullToken)
                    reader.Move(1)
                    Return New ValueExpress(tkn.GetToken(Of IToken)())

                Case GetType(LParenToken)
                    reader.Move(1)
                    Return CreateParenExpress(reader, Me.NextParser)

                Case GetType(PlusToken), GetType(MinusToken), GetType(NotToken)
                    reader.Move(1)
                    Dim nxtExper = Me.Parser(reader)
                    If TypeOf nxtExper Is ValueExpress Then
                        Return New UnaryExpress(tkn.GetToken(Of IToken)(), nxtExper)
                    Else
                        Throw New DSqlAnalysisException($"前置き演算子{tkn}が値の前に配置していません")
                    End If

                Case Else
                    Throw New DSqlAnalysisException("Factor要素の解析に失敗")
            End Select
        End Function

    End Class

#End Region

End Module
