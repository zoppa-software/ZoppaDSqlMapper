Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports ZoppaDSqlCompiler.Environments
Imports ZoppaDSqlCompiler.Express
Imports ZoppaDSqlCompiler.Tokens
Imports ZoppaDSqlCompiler.TokenCollection
Imports Microsoft.Extensions.Logging

''' <summary>トークンを解析する。</summary>
Friend Module ParserAnalysis

    ''' <summary>動的SQLの置き換えを実行します。</summary>
    ''' <param name="sqlQuery">動的SQL。</param>
    ''' <param name="parameter">パラメータ。</param>
    ''' <returns>置き換え結果。</returns>
    Public Function Replase(sqlQuery As String, Optional parameter As Object = Nothing) As String
        ' トークン配列に変換
        Dim tokens = LexicalAnalysis.Compile(sqlQuery)
        Logger.Value?.LogDebug("tokens")
        For i As Integer = 0 To tokens.Count - 1
            Logger.Value?.LogDebug($"{i + 1} : {tokens(i)}")
        Next

        ' 置き換え式を階層化する
        Dim hierarchy = CreateHierarchy(tokens)

        ' 置き換え式を解析して出力文字列を取得します。
        Dim buffer As New StringBuilder()
        Dim values = New EnvironmentObjectValue(parameter)
        Dim ansParts As New List(Of EvaParts)()
        Evaluation(hierarchy.Children, buffer, values, ansParts)

        ' 空白行を取り除いて返す
        Dim ans As New StringBuilder(buffer.Length)
        Using sr As New StringReader(buffer.ToString())
            Do While sr.Peek() <> -1
                Dim ln = sr.ReadLine()
                If ln.Trim() <> "" Then ans.AppendLine(ln.TrimEnd())
            Loop
        End Using
        Return ans.ToString().Trim()
    End Function

    ''' <summary>置き換え式の階層を作成します。</summary>
    ''' <param name="src">トークンリスト。</param>
    ''' <returns>階層リンク。</returns>
    Public Function CreateHierarchy(src As List(Of TokenPosition)) As TokenHierarchy
        Dim root As New TokenHierarchy(Nothing)
        If src?.Count > 0 Then
            ' 評価のため並び変え
            Dim tmp As New List(Of TokenPosition)(src)
            tmp.Reverse()

            ' 階層を作成
            CreateHierarchy("", root, tmp)
        End If
        Return root
    End Function

    ''' <summary>置き換え式の階層を作成します。</summary>
    ''' <param name="errMsg">エラーメッセージ。</param>
    ''' <param name="node">階層ノード。</param>
    ''' <param name="tokens">対象トークンリスト。</param>
    ''' <param name="partitionTokens"></param>
    Private Sub CreateHierarchy(errMsg As String,
                                    node As TokenHierarchy,
                                    tokens As List(Of TokenPosition),
                                    ParamArray partitionTokens As String())
        Dim limit As New HashSet(Of String)()
        For Each tkn In partitionTokens
            limit.Add(tkn)
        Next

        Do While tokens.Count > 0
            Dim tkn = tokens(tokens.Count - 1)
            If limit.Contains(tkn.TokenName) Then
                ' 末端トークンなら階層終了
                Return
            Else
                Dim cnode = node.AddChild(tkn)
                Select Case tkn.TokenName
                    Case NameOf(IfToken)
                        tokens.RemoveAt(tokens.Count - 1)
                        CreateHierarchy("ifブロックが閉じられていません", cnode, tokens, NameOf(ElseIfToken), NameOf(ElseToken), NameOf(EndIfToken))

                    Case NameOf(ElseIfToken)
                        tokens.RemoveAt(tokens.Count - 1)
                        CreateHierarchy("ifブロックが閉じられていません", cnode, tokens, NameOf(ElseToken), NameOf(EndIfToken))

                    Case NameOf(ElseToken)
                        tokens.RemoveAt(tokens.Count - 1)
                        CreateHierarchy("ifブロックが閉じられていません", cnode, tokens, NameOf(EndIfToken))

                    Case NameOf(ForEachToken)
                        tokens.RemoveAt(tokens.Count - 1)
                        CreateHierarchy("foreachブロックが閉じられていません", cnode, tokens, NameOf(EndForToken))

                    Case NameOf(TrimToken)
                        tokens.RemoveAt(tokens.Count - 1)
                        CreateHierarchy("trimブロックが閉じられていません", cnode, tokens, NameOf(EndTrimToken))

                    Case Else
                        tokens.RemoveAt(tokens.Count - 1)
                End Select
            End If
        Loop
        If partitionTokens.Count > 0 Then
            Throw New DSqlAnalysisException(errMsg)
        End If
    End Sub

    ''' <summary>トークンリストを評価します。</summary>
    ''' <param name="children">階層情報。</param>
    ''' <param name="buffer">出力先バッファ。</param>
    ''' <param name="prm">環境値情報。</param>
    ''' <param name="ansParts">評価結果リスト。</param>
    Private Sub Evaluation(children As List(Of TokenHierarchy),
                           buffer As StringBuilder,
                           prm As IEnvironmentValue,
                           ansParts As List(Of EvaParts))
        Dim tmp As New List(Of TokenHierarchy)(children)
        tmp.Reverse()

        Do While tmp.Count > 0
            Dim chd = tmp(tmp.Count - 1)

            Select Case chd.TargetToken.TokenName
                Case NameOf(IfToken)
                    EvaluationIf(tmp, buffer, prm, ansParts)

                Case NameOf(ForEachToken)
                    EvaluationFor(chd, buffer, prm, ansParts)
                    tmp.RemoveAt(tmp.Count - 1)

                Case NameOf(TrimToken)
                    EvaluationTrim(chd, buffer, prm, ansParts)
                    tmp.RemoveAt(tmp.Count - 1)

                Case NameOf(ReplaseToken)
                    Dim rtoken = chd.TargetToken.GetToken(Of ReplaseToken)()
                    If rtoken IsNot Nothing Then
                        Dim rval = prm.GetValue(If(rtoken.Contents?.ToString(), ""))
                        Dim ans = GetRefValue(rval, rtoken.IsEscape)
                        buffer.Append(ans)
                        ansParts.Add(New EvaParts(chd.TargetToken, ans.ToString()))
                    End If
                    tmp.RemoveAt(tmp.Count - 1)

                Case NameOf(EndIfToken)
                        ' ※ EvaluationIfで処理するため不要です

                Case NameOf(EndForToken), NameOf(EndTrimToken)
                    ansParts.Add(New EvaParts(chd.TargetToken, ""))
                    tmp.RemoveAt(tmp.Count - 1)

                Case Else
                    buffer.Append(chd.TargetToken.Contents)
                    ansParts.Add(New EvaParts(chd.TargetToken, chd.TargetToken.Contents.ToString()))
                    tmp.RemoveAt(tmp.Count - 1)
            End Select
        Loop
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

    ''' <summary>Ifを評価します。</summary>
    ''' <param name="tmp">階層情報。</param>
    ''' <param name="buffer">出力先バッファ。</param>
    ''' <param name="prm">環境値情報。</param>
    ''' <param name="ansParts">評価結果リスト。</param>
    Private Sub EvaluationIf(tmp As List(Of TokenHierarchy),
                                 buffer As StringBuilder,
                                 prm As IEnvironmentValue,
                                 ansParts As List(Of EvaParts))
        ' If、ElseIf、Elseブロックを集める
        Dim blocks As New List(Of TokenHierarchy)()
        Dim endifTkn As TokenPosition = Nothing
        For i As Integer = tmp.Count - 1 To 0 Step -1
            Dim iftkn = tmp(i)
            endifTkn = iftkn.TargetToken
            tmp.RemoveAt(tmp.Count - 1)
            If iftkn.TargetToken.TokenName <> NameOf(EndIfToken) Then
                blocks.Add(iftkn)
            Else
                Exit For
            End If
        Next

        ' If、ElseIf、Elseブロックを評価
        Dim output = False
        For Each iftkn In blocks
            Select Case iftkn.TargetToken.TokenName
                Case NameOf(IfToken), NameOf(ElseIfToken)
                    ' 条件を評価して真ならば、ブロックを出力
                    Dim ifans = Executes(iftkn.TargetToken.GetCommandToken().CommandTokens, prm)
                    If TypeOf ifans.Contents Is Boolean AndAlso CBool(ifans.Contents) Then
                        ansParts.Add(New EvaParts(iftkn.TargetToken, ""))
                        Evaluation(iftkn.Children, buffer, prm, ansParts)
                        output = True
                        Exit For
                    End If

                Case NameOf(ElseToken)
                    ansParts.Add(New EvaParts(iftkn.TargetToken, ""))
                    Evaluation(iftkn.Children, buffer, prm, ansParts)
                    output = True
                    Exit For
            End Select
        Next

        If output Then
            ' EndIfトークンを追加
            ansParts.Add(New EvaParts(endifTkn, ""))
        Else
            ' 非表示トークンを追加
            For Each iftkn In blocks
                ansParts.Add(New EvaParts(iftkn.TargetToken))
            Next
            ansParts.Add(New EvaParts(endifTkn))
        End If
    End Sub

    ''' <summary>Foreachを評価します。</summary>
    ''' <param name="fortoken">Foreachブロック。</param>
    ''' <param name="buffer">出力先バッファ。</param>
    ''' <param name="prm">環境値情報。</param>
    ''' <param name="ansParts">評価結果トークンリスト。</param>
    Private Sub EvaluationFor(forToken As TokenHierarchy,
                                  buffer As StringBuilder,
                                  prm As IEnvironmentValue,
                                  ansParts As List(Of EvaParts))
        Dim forTkn = forToken.TargetToken.GetToken(Of ForEachToken)()
        ansParts.Add(New EvaParts(forToken.TargetToken, ""))

        ' カウンタ変数
        Dim valkey As String = ""

        ' ループ元コレクション
        Dim collection As IEnumerable = Nothing

        ' 構文を解析して変数、ループ元コレクションを取得
        With forTkn
            If .CommandTokens.Count = 3 AndAlso
                   .CommandTokens(0).TokenName = NameOf(IdentToken) AndAlso
                   .CommandTokens(1).TokenName = NameOf(InToken) AndAlso
                   .CommandTokens(2).TokenName = NameOf(IdentToken) Then
                Dim colln = prm.GetValue(If(.CommandTokens(2).Contents?.ToString(), ""))
                If TypeOf colln Is IEnumerable Then
                    valkey = If(.CommandTokens(0).Contents?.ToString(), "")
                    collection = CType(colln, IEnumerable)
                End If
            End If
        End With

        ' Foreachして出力
        For Each v In collection
            prm.AddVariant(valkey, v)
            Evaluation(forToken.Children, buffer, prm, ansParts)
        Next

        prm.LocalVarClear()
    End Sub

    ''' <summary>Trimを評価します。</summary>
    ''' <param name="trimToken">Trimブロック。</param>
    ''' <param name="buffer">出力先バッファ。</param>
    ''' <param name="prm">環境値情報。</param>
    Private Sub EvaluationTrim(trimToken As TokenHierarchy,
                                   buffer As StringBuilder,
                                   prm As IEnvironmentValue,
                                   ansParts As List(Of EvaParts))
        Dim answer As New StringBuilder()

        ' Trim内のトークンを評価
        Dim tmpbuf As New StringBuilder()
        Dim subAnsParts As New List(Of EvaParts)()
        Evaluation(trimToken.Children, tmpbuf, prm, subAnsParts)

        ' 両端の空白を取り除く
        BothTrim(subAnsParts)

        ' トリムルールに従ってトリムします
        '
        ' 1. foreachルール
        ' 2. whereルール
        ' 3. setルール
        ' 4. 通常ルール
        If subAnsParts.First().TokenPos.TokenName = NameOf(ForEachToken) AndAlso
               subAnsParts.Last().TokenPos.TokenName = NameOf(EndForToken) Then             ' 1
            RemoveTrimEndByForeach(trimToken.TargetToken.GetToken(Of TrimToken)(), subAnsParts, ",")
            For Each parts In subAnsParts
                ansParts.Add(parts)
                buffer.Append(parts.OutString)
            Next

        ElseIf subAnsParts.First().TokenPos.TokenName = NameOf(QueryToken) AndAlso
                   subAnsParts.First().OutString?.Trim().ToLower().StartsWith("where") Then '2
            buffer.Append(RemoveAndOrByWhere(subAnsParts))

        ElseIf subAnsParts.First().TokenPos.TokenName = NameOf(QueryToken) AndAlso
                   subAnsParts.First().OutString?.Trim().ToLower().StartsWith("set") Then   ' 3
            'RemoveAndOrByWhere(subAnsParts)
            'RemoveParentByWhere(subAnsParts)

            'Dim tmp As New StringBuilder()
            'For Each tkn In subAnsParts
            '    If tkn.isOutpit Then
            '        ansParts.Add(tkn)
            '        tmp.Append(tkn.outString)
            '    End If
            'Next

            'Dim tmpstr = tmp.ToString()
            'If tmpstr.Trim().ToLower() <> "where" Then
            '    buffer.Append(tmpstr)
            'End If

        Else    ' 4
            RemoveTrimEndByForeach(trimToken.TargetToken.GetToken(Of TrimToken)(), subAnsParts, "")
            For Each tkn In subAnsParts
                ansParts.Add(tkn)
                buffer.Append(tkn.OutString)
            Next
        End If
    End Sub

    ''' <summary>両端の空白をトリムします。</summary>
    ''' <param name="subAnsParts">評価部分リスト。</param>
    Private Sub BothTrim(subAnsParts As List(Of EvaParts))
        ' 先頭の空白を取り除く
        Do While subAnsParts.Count > 0
            If subAnsParts(0).OutString?.Trim() <> "" OrElse (subAnsParts(0).IsControlToken AndAlso subAnsParts(0).IsOutpit) Then
                Exit Do
            Else
                subAnsParts.RemoveAt(0)
            End If
        Loop

        ' 末尾の空白を取り除く
        For i As Integer = subAnsParts.Count - 1 To 0 Step -1
            If subAnsParts(i).OutString?.Trim() <> "" OrElse (subAnsParts(i).IsControlToken AndAlso subAnsParts(i).IsOutpit) Then
                Exit For
            Else
                subAnsParts.RemoveAt(i)
            End If
        Next
    End Sub

    ''' <summary>末尾の指定文字列をトリムします。</summary>
    ''' <param name="trimTkn">Trimトークン。</param>
    ''' <param name="ansParts">評価部分リスト。</param>
    ''' <param name="defTrmStr">デフォルトトリム文字列。</param>
    Private Sub RemoveTrimEndByForeach(trimTkn As TrimToken, ansParts As List(Of EvaParts), defTrmStr As String)
        ' トリムする文字列を取得
        Dim srctrm = trimTkn.TrimString?.Trim()
        Dim trmstr = If(srctrm = "", defTrmStr, srctrm)

        If trmstr <> "" Then
            ' 末尾の出力トークンの末尾を削除
            For i As Integer = ansParts.Count - 1 To 0 Step -1
                If ansParts(i).TokenPos.TokenName = NameOf(QueryToken) AndAlso Not ansParts(i).IsSpace Then
                    Dim str = ansParts(i).OutString.TrimEnd()
                    If str.EndsWith(trmstr) Then
                        ansParts(i).OutString = str.Substring(0, str.Length - trmstr.Length)
                    End If
                    Exit For
                End If
            Next
        End If
    End Sub

    ''' <summary>where句のためのトリム処理を行います。</summary>
    ''' <param name="ansParts">評価部分リスト。</param>
    Private Function RemoveAndOrByWhere(ansParts As List(Of EvaParts)) As String
        Dim copyParts As New List(Of EvaParts)()
        Dim partsLink As New Dictionary(Of EvaParts, List(Of EvaParts))()
        For Each ps In ansParts
            If ps.TokenPos.TokenName = NameOf(QueryToken) Then
                ' 出力トークンならば and/or を評価するため構文解析して評価

                ' 1. トークンの分割
                Dim tokens = CreateTokenTypes(ps)

                ' 2. グループ化するためにカウント
                Dim def = CountPartsDefine(tokens)

                ' 3. グループ化した単位で評価部分に変換して、階層化して保持
                Dim children As New List(Of EvaParts)()
                Dim i As Integer = 0
                Do While i < tokens.Count
                    Dim buf As New StringBuilder()
                    buf.Append(tokens(i).str)
                    i += 1
                    For j As Integer = i To tokens.Count - 1
                        If def(j - 1) = def(j) Then
                            buf.Append(tokens(j).str)
                            i += 1
                        Else
                            Exit For
                        End If
                    Next
                    children.Add(New EvaParts(ps.TokenPos, buf.ToString()))
                Loop
                partsLink.Add(ps, children)

                ' 4. 評価用リストに追加
                copyParts.AddRange(children)
            Else
                ' 出力トークン以外はそのまま評価
                copyParts.Add(ps)
            End If
        Next

        ' 評価リストのポインタを生成
        Dim parts = New EvaPartsPointer(copyParts)

        ' 先頭の where を削除
        parts.SkipWord("where")

        ' 評価を行い、出力する文字列を作成
        LogicalTrim(parts)
        Dim tmp As New StringBuilder()
        For Each tkn In parts
            If tkn.IsOutpit Then
                ansParts.Add(tkn)
                tmp.Append(tkn.OutString)
            End If
        Next

        Dim tmpstr = tmp.ToString()
        If tmpstr.Trim().ToLower() <> "where" Then
            ' 出力する文字列があれば表示制御
            For Each parent In partsLink.Keys
                Dim buf As New StringBuilder()
                Dim view = False
                For Each ch In partsLink(parent)
                    If ch.IsOutpit Then
                        buf.Append(ch.OutString)
                        view = True
                    End If
                Next
                parent.OutString = buf.ToString()
                parent.IsOutpit = view
            Next
            Return tmpstr
        Else
            ' 出力する文字列がなければ全て非表示
            For Each ps In ansParts
                ps.IsOutpit = False
            Next
            Return ""
        End If
    End Function

    ''' <summary>トークンの情報リストを作成します。</summary>
    ''' <param name="ps">評価部分。</param>
    ''' <returns>トークンの情報リスト。</returns>
    Private Function CreateTokenTypes(ps As EvaParts) As List(Of (str As String, keywd As Boolean, space As Boolean))
        Dim tokens As New List(Of (str As String, keywd As Boolean, space As Boolean))()

        Dim ch = ps.OutString.ToArray()
        Dim i As Integer = 0
        Do While i < ch.Length
            If ch(i) = "("c Then
                tokens.Add(("(", True, False))
                i += 1
            ElseIf ch(i) = ")"c Then
                tokens.Add((")", True, False))
                i += 1
            ElseIf Char.IsWhiteSpace(ch(i)) Then
                Dim buf As New StringBuilder()
                For j As Integer = i To ch.Length - 1
                    If Char.IsWhiteSpace(ch(j)) Then
                        buf.Append(ch(j))
                        i += 1
                    Else
                        Exit For
                    End If
                Next
                tokens.Add((buf.ToString(), False, True))
            Else
                Dim buf As New StringBuilder()
                For j As Integer = i To ch.Length - 1
                    If Not Char.IsWhiteSpace(ch(j)) OrElse ch(j) = "("c OrElse ch(j) = ")"c Then
                        buf.Append(ch(j))
                        i += 1
                    Else
                        Exit For
                    End If
                Next
                Dim str = buf.ToString()
                tokens.Add((str, str.ToLower() = "and" OrElse str.ToLower() = "or", False))
            End If
        Loop

        Return tokens
    End Function

    ''' <summary>(、)、and、or、それ以外を区別できるようにカウントする。</summary>
    ''' <param name="tokens">トークンの情報リスト。</param>
    ''' <returns>区別用カウント。</returns>
    Private Function CountPartsDefine(tokens As List(Of (str As String, keywd As Boolean, space As Boolean))) As Integer()
        Dim def = New Integer(tokens.Count - 1) {}

        Dim cnt As Integer = 0
        For j As Integer = 0 To tokens.Count - 1
            If tokens(j).keywd Then
                ' キーワードは別定義としてカウントアップ
                If j > 0 AndAlso Not tokens(j - 1).keywd Then cnt += 1
                def(j) = cnt
                cnt += 1

            ElseIf tokens(j).space AndAlso
                    (j = 0 OrElse tokens(j - 1).keywd) AndAlso
                    (j = tokens.Count - 1 OrElse tokens(j + 1).keywd) Then
                ' 空白かつ左右が別トークンなら別グループ
                def(j) = cnt
                cnt += 1
            Else
                def(j) = cnt
            End If
        Next

        Return def
    End Function

    ''' <summary>トリムのための and/orを評価します。</summary>
    ''' <param name="parts">評価部分ポインタ。</param>
    ''' <returns>表示があるならば真。</returns>
    Private Function LogicalTrim(parts As EvaPartsPointer) As Boolean
        ' 左辺要素を取得
        Dim tml = FactorTrim(parts)

        ' and/orを見つけるまで要素を進める
        Do While parts.HasNext
            Dim ope = parts.Current
            If ope.OutString.ToLower() = "and" OrElse
                   ope.OutString.ToLower() = "or" Then
                parts.Increment()

                ' 右辺要素を取得
                Dim tmr = FactorTrim(parts)

                ' 右辺、左辺のどちらかが非表示ならば and/orを非表示
                If Not tml OrElse Not tmr Then
                    ope.IsOutpit = False
                End If

                ' 右辺、左辺の表示結果を orする
                tml = tml Or tmr
            Else
                Exit Do
            End If
        Loop
        Return tml
    End Function

    ''' <summary>トリムのための要素を評価します。</summary>
    ''' <param name="parts">評価部分ポインタ。</param>
    ''' <returns>表示があるならば真。</returns>
    Private Function FactorTrim(parts As EvaPartsPointer) As Boolean
        Dim facts As New List(Of EvaParts)()
        Dim isOut As Boolean = False

        ' (、and、orの要素を取得するまでリストに保持
        Do While parts.HasNext
            If parts.Current.OutString = "(" Then
                isOut = isOut Or ParenTrim(parts)
            ElseIf parts.Current.OutString = ")" Then
                Exit Do
            ElseIf parts.Current.OutString.ToLower() = "and" OrElse
                       parts.Current.OutString.ToLower() = "or" Then
                Exit Do
            Else
                facts.Add(parts.Current)
                parts.Increment()
            End If
        Loop

        If isOut Then
            ' 要素を出力するので真
            Return True
        Else
            ' 出力する要素があるならば真、全て非表示ならば偽
            For Each pt In facts
                If Not pt.IsSpace AndAlso pt.IsOutpit Then
                    Return True
                End If
            Next
            For Each pt In facts
                pt.IsOutpit = False
            Next
            Return False
        End If
    End Function

    ''' <summary>トリムのための括弧を評価します。</summary>
    ''' <param name="parts">評価部分ポインタ。</param>
    ''' <returns>表示があるならば真。</returns>
    Private Function ParenTrim(parts As EvaPartsPointer) As Boolean
        ' (を取得
        Dim lParen = parts.Current
        parts.Increment()
        Dim fptr = parts.Index

        ' and/orをトリム
        Dim isOut = LogicalTrim(parts)

        ' )を取得
        Dim rParen = parts.Current
        Dim rptr = parts.Index - 1
        parts.Increment()

        ' 表示するなら真、表示しないならば()を非表示にして偽
        If isOut Then
            For i As Integer = fptr To parts.Count - 1
                With parts(i)
                    If .IsOutpit Then
                        If .IsSpace Then
                            .IsOutpit = False
                        Else
                            Exit For
                        End If
                    End If
                End With
            Next
            For i As Integer = rptr To 0 Step -1
                With parts(i)
                    If .IsOutpit Then
                        If .IsSpace Then
                            .IsOutpit = False
                        Else
                            Exit For
                        End If
                    End If
                End With
            Next
            Return True
        Else
            lParen.IsOutpit = False
            rParen.IsOutpit = False
            Return False
        End If
    End Function

    ''' <summary>式を解析して結果を取得します。</summary>
    ''' <param name="expression">式文字列。</param>
    ''' <param name="parameter">パラメータ。</param>
    ''' <returns>解析結果。</returns>
    Public Function Executes(expression As String, Optional parameter As Object = Nothing) As IToken
        ' トークン配列に変換
        Dim tokens = LexicalAnalysis.SplitToken(expression)

        ' 式木を実行
        Return Executes(tokens, New EnvironmentObjectValue(parameter))
    End Function

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

            Select Case tkn.TokenName
                Case NameOf(LParenToken)
                    tmp.Add(tkn)
                    lv += 1

                Case NameOf(RParenToken)
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
            If tkn.TokenName = NameOf(LParenToken) Then
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
                Select Case ope.TokenName
                    Case NameOf(AndToken)
                        reader.Move(1)
                        tml = New AndExpress(tml, Me.NextParser.Parser(reader))

                    Case NameOf(OrToken)
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
                Select Case ope.TokenName
                    Case NameOf(EqualToken)
                        reader.Move(1)
                        tml = New EqualExpress(tml, Me.NextParser.Parser(reader))

                    Case NameOf(NotEqualToken)
                        reader.Move(1)
                        tml = New NotEqualExpress(tml, Me.NextParser.Parser(reader))

                    Case NameOf(GreaterToken)
                        reader.Move(1)
                        tml = New GreaterExpress(tml, Me.NextParser.Parser(reader))

                    Case NameOf(GreaterEqualToken)
                        reader.Move(1)
                        tml = New GreaterEqualExpress(tml, Me.NextParser.Parser(reader))

                    Case NameOf(LessToken)
                        reader.Move(1)
                        tml = New LessExpress(tml, Me.NextParser.Parser(reader))

                    Case NameOf(LessEqualToken)
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
                Select Case ope.TokenName
                    Case NameOf(PlusToken)
                        reader.Move(1)
                        tml = New PlusExpress(tml, Me.NextParser.Parser(reader))

                    Case NameOf(MinusToken)
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
                Select Case ope.TokenName
                    Case NameOf(MultiToken)
                        reader.Move(1)
                        tml = New MultiExpress(tml, Me.NextParser.Parser(reader))

                    Case NameOf(DivToken)
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

            Select Case tkn.TokenName
                Case NameOf(IdentToken), NameOf(NumberToken), NameOf(StringToken),
                         NameOf(QueryToken), NameOf(ReplaseToken), NameOf(ObjectToken),
                         NameOf(TrueToken), NameOf(FalseToken), NameOf(NullToken)
                    reader.Move(1)
                    Return New ValueExpress(tkn.GetToken(Of IToken)())

                Case NameOf(LParenToken)
                    reader.Move(1)
                    Return CreateParenExpress(reader, Me.NextParser)

                Case NameOf(PlusToken), NameOf(MinusToken), NameOf(NotToken)
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

End Module
