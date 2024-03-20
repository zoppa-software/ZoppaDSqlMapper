Option Strict On
Option Explicit On

Imports System.Text
Imports ZoppaDSqlReplace.Environments
Imports ZoppaDSqlReplace.Express
Imports ZoppaDSqlReplace.Tokens
Imports ZoppaDSqlReplace.TokenCollection
Imports System.Security.Cryptography
Imports System.Data.Common

''' <summary>トークンを解析する。</summary>
Friend Module ParserAnalysis

    ' 三項演算子パーサ
    Private ReadOnly mMultiParser As New MultiEvalParser()

    ' 論理演算子パーサ
    Private ReadOnly mLogicalParser As New LogicalParser()

    ' 比較演算子パーサ
    Private ReadOnly mCompParser As New ComparisonParser()

    ' 加減算パーサ
    Private ReadOnly mAddOrSubParser As New AddOrSubParser()

    ' 乗除算パーサ
    Private ReadOnly mMultiOrDivParser As New MultiOrDivParser()

    ' 参照パーサ
    Private ReadOnly mRefParser As New ReferenceParser()

    ' 要素パーサ
    Private ReadOnly mFacParser As New FactorParser()

    ' 括弧パーサ
    Private ReadOnly mParenParser As New ParenParser()

    ''' <summary>コンストラクタ。</summary>
    Sub New()
        mMultiParser.NextParser = mLogicalParser
        mLogicalParser.NextParser = mCompParser
        mCompParser.NextParser = mAddOrSubParser
        mAddOrSubParser.NextParser = mMultiOrDivParser
        mMultiOrDivParser.NextParser = mRefParser
        mRefParser.NextParser = mFacParser
        mFacParser.NextParser = mParenParser
        mParenParser.NextParser = mMultiParser
    End Sub

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
        Dim lines = buffer.ToString().Split(New String() {Environment.NewLine, vbLf}, StringSplitOptions.None)

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
                    Dim ifTokens = CollectBlockToken(sqlQuery, reader, New Type() {GetType(IfToken)}, GetType(EndIfToken))
                    EvaluationIf(sqlQuery, tkn.Position, tkn.GetToken(Of IfToken)(), ifTokens, buffer, parameter)

                Case GetType(ElseIfToken), GetType(ElseToken), GetType(EndIfToken)
                    Throw New DSqlAnalysisException($"ifが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")

                Case GetType(ForToken)
                    reader.Move(1)
                    Dim forTokens = CollectBlockToken(sqlQuery, reader, New Type() {GetType(ForToken), GetType(ForEachToken)}, GetType(EndForToken))
                    EvaluationFor(sqlQuery, tkn.Position, tkn.GetToken(Of ForToken)(), forTokens, buffer, parameter)

                Case GetType(ForEachToken)
                    reader.Move(1)
                    Dim forTokens = CollectBlockToken(sqlQuery, reader, New Type() {GetType(ForToken), GetType(ForEachToken)}, GetType(EndForToken))
                    EvaluationForEach(sqlQuery, tkn.Position, tkn.GetToken(Of ForEachToken)(), forTokens, buffer, parameter)

                Case GetType(EndForToken)
                    Throw New DSqlAnalysisException($"forが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")

                Case GetType(SelectToken)
                    reader.Move(1)
                    Dim selTokens = CollectBlockToken(sqlQuery, reader, New Type() {GetType(SelectToken)}, GetType(EndSelectToken))
                    EvaluationSelect(sqlQuery, tkn.GetToken(Of SelectToken)(), selTokens, buffer, parameter)

                Case GetType(CaseToken), GetType(EndSelectToken)
                    Throw New DSqlAnalysisException($"selectが開始されていません。{vbCrLf}{sqlQuery}:{tkn.Position}")

                Case GetType(TrimToken)
                    reader.Move(1)
                    Dim trimTokens = CollectBlockToken(sqlQuery, reader, New Type() {GetType(TrimToken)}, GetType(EndTrimToken))
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
                                       startTokenName As Type(),
                                       endTokenName As Type) As List(Of TokenPosition)
        Dim res As New List(Of TokenPosition)()
        Dim startToken = reader.Current

        Dim nest As Integer = 0
        Do While reader.HasNext
            Dim tkn = reader.Current
            If startTokenName.Any(Function(t) tkn.TokenType Is t) Then
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
    ''' <param name="tempPos">評価位置。</param>
    ''' <param name="sifToken">Ifのトークンリスト。</param>
    ''' <param name="tokens">ブロック内のトークンリスト。</param>
    ''' <param name="buffer">結果バッファ。</param>
    ''' <param name="parameter">パラメータ。</param>
    Private Sub EvaluationIf(sqlQuery As String,
                             tempPos As Integer,
                             sifToken As IToken,
                             tokens As List(Of TokenPosition),
                             buffer As StringBuilder,
                             parameter As IEnvironmentValue)
        Dim blocks As New List(Of (condition As IToken, block As List(Of TokenPosition))) From {
            (sifToken, New List(Of TokenPosition)())
        }

        '----------------
        ' トークン解析
        '----------------
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

        '----------------
        ' 条件分岐実行
        '----------------
        Try
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
        Catch ex As Exception
            Throw New DSqlAnalysisException($"ifの構文を間違えています。{vbCrLf}{sqlQuery}:{tempPos}", ex)
        End Try
    End Sub

    ''' <summary>Forを評価します。</summary>
    ''' <param name="sqlQuery">元のSQL。</param>
    ''' <param name="tempPos">評価位置。</param>
    ''' <param name="sforToken">Forのトークンリスト。</param>
    ''' <param name="tokens">ブロック内のトークンリスト。</param>
    ''' <param name="buffer">結果バッファ。</param>
    ''' <param name="parameter">パラメータ。</param>
    Private Sub EvaluationFor(sqlQuery As String,
                              tempPos As Integer,
                              sforToken As ForToken,
                              tokens As List(Of TokenPosition),
                              buffer As StringBuilder,
                              parameter As IEnvironmentValue)
        '----------------
        ' トークン解析
        '----------------
        Dim ts As New TokenStream(CType(sforToken.Contents, List(Of TokenPosition)))

        ' カウンタ変数を取得
        Dim varName As String
        Dim varContents = ts.Current.GetToken(Of IdentToken)().Contents
        If varContents IsNot Nothing AndAlso Not parameter.IsDefainedName(varContents.ToString()) Then
            varName = varContents.ToString()
        Else
            Throw New DSqlAnalysisException($"変数が定義済みです。{vbCrLf}{sqlQuery}:{tempPos}")
        End If
        ts.Move(1)

        ' = トークンを取得
        If ts.Current.TokenType IsNot GetType(EqualToken) Then
            Throw New DSqlAnalysisException($"forの構文を間違えています。{vbCrLf}{sqlQuery}:{tempPos}")
        End If
        ts.Move(1)

        ' 開始値を取得
        Dim startToken = mMultiParser.Parser(ts)

        ' toトークンを取得
        Dim toToken = ts.Current.GetToken(Of IdentToken)()
        If toToken Is Nothing OrElse toToken.ToString().ToLower() <> "to" Then
            Throw New DSqlAnalysisException($"forの構文を間違えています。{vbCrLf}{sqlQuery}:{tempPos}")
        End If
        ts.Move(1)

        ' 終了値を取得
        Dim endToken = mMultiParser.Parser(ts)

        If ts.HasNext Then
            Throw New DSqlAnalysisException($"forの構文を間違えています。{vbCrLf}{sqlQuery}:{tempPos}")
        End If

        '----------------
        ' ループ実行
        '----------------
        Try
            For i As Integer = Convert.ToInt32(startToken.Executes(parameter).Contents) To Convert.ToInt32(endToken.Executes(parameter).Contents)
                parameter.AddVariant(varName, i)

                Dim lclbuf As New StringBuilder()
                ReplaseQuery(sqlQuery, New TokenStream(tokens), parameter, lclbuf)
                buffer.Append(lclbuf)

                parameter.RemoveVariant(varName)
            Next
        Catch ex As Exception
            Throw New DSqlAnalysisException($"forの構文を間違えています。{vbCrLf}{sqlQuery}:{tempPos}", ex)
        End Try
    End Sub

    ''' <summary>ForEachを評価します。</summary>
    ''' <param name="templateStr">テンプレート。</param>
    ''' <param name="tempPos">評価位置。</param>
    ''' <param name="sforToken">Forのトークンリスト。</param>
    ''' <param name="tokens">ブロック内のトークンリスト。</param>
    ''' <param name="buffer">結果バッファ。</param>
    ''' <param name="parameter">パラメータ。</param>
    Private Sub EvaluationForEach(templateStr As String,
                                      tempPos As Integer,
                                      sforToken As ForEachToken,
                                      tokens As List(Of TokenPosition),
                                      buffer As StringBuilder,
                                      parameter As IEnvironmentValue)
        '----------------
        ' トークン解析
        '----------------
        Dim ts As New TokenStream(CType(sforToken.Contents, List(Of TokenPosition)))

        ' カウンタ変数を取得
        Dim varName As String
        Dim varContents = ts.Current.GetToken(Of IdentToken)().Contents
        If varContents IsNot Nothing AndAlso Not parameter.IsDefainedName(varContents.ToString()) Then
            varName = varContents.ToString()
        Else
            Throw New DSqlAnalysisException($"変数が定義済みです。{vbCrLf}{templateStr}:{tempPos}")
        End If
        ts.Move(1)

        ' in、:トークンを取得
        Dim toToken = ts.Current.GetToken(Of InToken)()
        If toToken Is Nothing OrElse toToken.ToString().ToLower() <> "in" Then
            Throw New DSqlAnalysisException($"foreachの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}")
        End If
        ts.Move(1)

        ' コレクションを取得
        Dim collectionToken = mMultiParser.Parser(ts)
        Dim collection = TryCast(collectionToken.Executes(parameter).Contents, IEnumerable)
        If collection Is Nothing Then
            Throw New DSqlAnalysisException($"foreachの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}")
        End If

        If ts.HasNext Then
            Throw New DSqlAnalysisException($"foreachの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}")
        End If

        '----------------
        ' ループ実行
        '----------------
        Try
            For Each v In collection
                parameter.AddVariant(varName, v)

                Dim lclbuf As New StringBuilder()
                ReplaseQuery(templateStr, New TokenStream(tokens), parameter, lclbuf)
                buffer.Append(lclbuf)

                parameter.RemoveVariant(varName)
            Next
        Catch ex As Exception
            Throw New DSqlAnalysisException($"foreachの構文を間違えています。{vbCrLf}{templateStr}:{tempPos}", ex)
        End Try
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

        For Each twd In strimToken.TrimStrings
            If tarStr.TrimEnd().EndsWith(twd.Word, StringComparison.CurrentCultureIgnoreCase) Then
                Dim idx = tarStr.LastIndexOf(twd.Word, StringComparison.CurrentCultureIgnoreCase)
                If Not twd.IsKey OrElse (idx = 0 OrElse Char.IsWhiteSpace(tarStr(idx - 1))) Then
                    tarStr = tarStr.Substring(0, idx).TrimEnd(" "c, vbTab(0))
                    Exit For
                End If
            End If
        Next

        If strimToken.IsBoth Then
            For Each twd In strimToken.TrimStrings
                If tarStr.TrimStart().StartsWith(twd.Word, StringComparison.CurrentCultureIgnoreCase) Then
                    Dim startIdx = tarStr.IndexOf(twd.Word, StringComparison.CurrentCultureIgnoreCase)
                    Dim endIdx = startIdx + twd.Word.Length
                    If Not twd.IsKey OrElse (endIdx = tarStr.Length OrElse Char.IsWhiteSpace(tarStr(endIdx))) Then
                        tarStr = tarStr.Remove(startIdx, twd.Word.Length).TrimStart(" "c, vbTab(0))
                        Exit For
                    End If
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
                    If EqualExpress.ExpressionEqual(caseans, caseVal) Then
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
        ' トークン解析
        Dim tknPtr = New TokenStream(tokens)
        Dim expr = mMultiParser.Parser(tknPtr)

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
    Private Function CreateParenExpress(Of TLParen, TRParen)(reader As TokenStream, nxtParser As IParser) As ParenExpress
        Dim tmp As New List(Of TokenPosition)()
        Dim lv As Integer = 0
        Do While reader.HasNext
            Dim tkn = reader.Current
            reader.Move(1)

            Select Case tkn.TokenType
                Case GetType(TLParen)
                    tmp.Add(tkn)
                    lv += 1

                Case GetType(TRParen)
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
                Return CreateParenExpress(Of LParenToken, RParenToken)(reader, Me.NextParser)
            Else
                Return Me.NextParser.Parser(reader)
            End If
        End Function

    End Class

    ''' <summary>カンマ分割解析。</summary>
    Private Class CommaSplitParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim res As New List(Of IExpression)()

            Do While reader.HasNext
                res.Add(Me.NextParser.Parser(reader))

                If reader.Current.TokenType = GetType(CommaToken) Then
                    reader.Move(1)
                Else
                    Exit Do
                End If
            Loop

            Return New CommaSplitExpress(res)
        End Function

    End Class

    ''' <summary>三項演算子解析。</summary>
    Private Class MultiEvalParser
        Implements IParser

        ''' <summary>次のパーサーを設定、取得する。</summary>
        Friend Property NextParser() As IParser

        ''' <summary>解析を実行する。</summary>
        ''' <param name="reader">入力トークンストリーム。</param>
        ''' <returns>解析結果。</returns>
        Public Function Parser(reader As TokenStream) As IExpression Implements IParser.Parser
            Dim tml = Me.NextParser.Parser(reader)

            If reader.HasNext AndAlso reader.Current.TokenType = GetType(QuestionToken) Then
                reader.Move(1)
                Dim tmTrue = Me.NextParser.Parser(reader)
                If reader.HasNext AndAlso reader.Current.TokenType = GetType(ColonToken) Then
                    reader.Move(1)
                    tml = New MultiEvalExpress(tml, tmTrue, Me.NextParser.Parser(reader))
                End If
            End If

            Return tml
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

    ''' <summary>参照解析。</summary>
    Private NotInheritable Class ReferenceParser
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
                    Case GetType(PeriodToken)
                        reader.Move(1)
                        tml = New ReferenceExpress(tml, Me.NextParser.Parser(reader))

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
                Case GetType(LParenToken)
                    reader.Move(1)
                    Return CreateParenExpress(Of LParenToken, RParenToken)(reader, Me.NextParser)

                Case GetType(LBracketToken)
                    reader.Move(1)
                    Dim arrParser As New CommaSplitParser With {.NextParser = Me.NextParser}
                    Return CreateParenExpress(Of LBracketToken, RBracketToken)(reader, arrParser)

                Case GetType(IdentToken)
                    reader.Move(1)
                    If reader.Current.TokenType Is GetType(LBracketToken) Then
                        reader.Move(1)
                        Dim numTkn = CreateParenExpress(Of LBracketToken, RBracketToken)(reader, Me.NextParser)
                        Return New ArrayValueExpress(tkn.GetToken(Of IdentToken)(), numTkn)
                    ElseIf reader.Current.TokenType Is GetType(LParenToken) Then
                        reader.Move(1)
                        Dim arrParser As New CommaSplitParser With {.NextParser = Me.NextParser}
                        Dim argTkn = CreateParenExpress(Of LParenToken, RParenToken)(reader, arrParser)
                        Return New MethodEvalExpress(tkn.GetToken(Of IdentToken)(), argTkn)
                    Else
                        Return New ValueExpress(tkn.GetToken(Of IdentToken)())
                    End If

                Case GetType(NumberToken), GetType(StringToken),
                         GetType(TrueToken), GetType(FalseToken),
                         GetType(NullToken), GetType(ObjectToken)
                    '     GetType(QueryToken), GetType(ReplaseToken),
                    reader.Move(1)
                    Return New ValueExpress(tkn.GetToken(Of IToken)())

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
