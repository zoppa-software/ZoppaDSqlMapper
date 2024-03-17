Option Strict On
Option Explicit On

Imports ZoppaDSqlReplace.TokenCollection

Namespace Tokens

    ''' <summary>Trimトークン。</summary>
    Public NotInheritable Class TrimToken
        Implements IToken, IControlToken

        ''' <summary>末尾からトリムする文字列を返します。</summary>
        Public ReadOnly Property TrimStrings As TrimWord()

        ''' <summary>両方トリムするかどうかを取得します。</summary>
        Public ReadOnly Property IsBoth As Boolean

        ''' <summary>Whereトリムするかどうかを取得します。</summary>
        Public ReadOnly Property IsWhere As Boolean

        ''' <summary>格納されている値を取得する。</summary>
        ''' <returns>格納値。</returns>
        Public ReadOnly Property Contents As Object Implements IToken.Contents
            Get
                Throw New NotImplementedException("使用できません")
            End Get
        End Property

        ''' <summary>トークン型を取得する。</summary>
        ''' <returns>トークン型。</returns>
        Public ReadOnly Property TokenType As Type Implements IToken.TokenType
            Get
                Return GetType(TrimToken)
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

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="isBoth">両方トリムするかどうか。</param>
        ''' <param name="isWTrim">Whereトリムを行うならば。</param>
        Public Sub New(isBoth As Boolean, isWTrim As Boolean)
            Me.TrimStrings = New TrimWord() {
                New TrimWord("where", True),
                New TrimWord(",", False),
                New TrimWord("and", True),
                New TrimWord("or", True),
                New TrimWord("()", False)
            }
            Me.IsBoth = isBoth
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="isBoth">両方トリムするかどうか。</param>
        ''' <param name="isWTrim">Whereトリムを行うならば。</param>
        ''' <param name="trimStr">末尾からトリム文字。</param>
        Public Sub New(isBoth As Boolean, isWTrim As Boolean, trimStr As String)
            Dim tokens = LexicalAnalysis.SplitToken(trimStr)
            Dim res As New List(Of IToken)()

            Dim pointer As New TokenStream(tokens)
            Dim nexted = False
            Do While pointer.HasNext
                Select Case pointer.Current.TokenType
                    Case GetType(StringToken)
                        If Not nexted Then
                            res.Add(pointer.Current.GetToken(Of StringToken))
                            pointer.Move(1)
                            nexted = True
                        Else
                            Throw New DSqlAnalysisException("トリム文字列に不正なトークンが含まれています。")
                        End If

                    Case GetType(NumberToken)
                        If Not nexted Then
                            res.Add(pointer.Current.GetToken(Of NumberToken))
                            pointer.Move(1)
                            nexted = True
                        Else
                            Throw New DSqlAnalysisException("トリム文字列に不正なトークンが含まれています。")
                        End If

                    Case GetType(CommaToken)
                        If nexted Then
                            pointer.Move(1)
                            nexted = False
                        Else
                            Throw New DSqlAnalysisException("トリム文字列に不正なトークンが含まれています。")
                        End If

                    Case Else
                        Throw New DSqlAnalysisException("トリム文字列に不正なトークンが含まれています。")
                End Select
            Loop

            Me.TrimStrings = res.Select(
                Function(s)
                    Dim wd = s.Contents.ToString()
                    Dim isKey = If(TryCast(s, StringToken)?.BracketChar = """"c, False)
                    Return New TrimWord(wd, isKey)
                End Function).ToArray()
            Me.IsBoth = isBoth
        End Sub

        ''' <summary>文字列条件を取得します。</summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            Return "Trim"
        End Function

        Public Structure TrimWord

            Public ReadOnly Word As String

            Public ReadOnly IsKey As Boolean

            Public Sub New(word As String, isKey As Boolean)
                Me.Word = word
                Me.IsKey = isKey
            End Sub

        End Structure

    End Class

End Namespace
