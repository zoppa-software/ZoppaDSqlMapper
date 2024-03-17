Option Strict On
Option Explicit On

Imports System.Linq.Expressions
Imports System.Reflection
Imports ZoppaDSqlReplace
Imports ZoppaDSqlReplace.Tokens
Imports Microsoft.Extensions.Logging

Namespace Environments

    ''' <summary>環境値情報（オブジェクト）。</summary>
    Public NotInheritable Class EnvironmentObjectValue
        Implements IEnvironmentValue

        ' パラメータの型
        Private ReadOnly mType As Type

        ' パラメータ実体
        Private ReadOnly mTarget As Object

        ' プロパティディクショナリ
        Private ReadOnly mPropDic As New Dictionary(Of String, PropertyInfo)

        ' ローカル変数
        Private ReadOnly mVariants As New Dictionary(Of String, Object)

        ' リトライフラグ
        Private _retry As Boolean = False

        ' リトライ中ならば真
        Private Property Retrying As Boolean
            Get
                SyncLock Me
                    Return Me._retry
                End SyncLock
            End Get
            Set(value As Boolean)
                SyncLock Me
                    Me._retry = value
                End SyncLock
            End Set
        End Property

        ''' <summary>コピーコンストラクタ。</summary>
        ''' <param name="other">コピー元。</param>
        Private Sub New(other As EnvironmentObjectValue)
            Me.mType = other.mType
            Me.mTarget = other.mTarget
            Me.mPropDic = other.mPropDic
            Me.mVariants = New Dictionary(Of String, Object)(other.mVariants)
            Me.Retrying = other.Retrying
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="type">パラメータの型。</param>
        ''' <param name="target">パラメータ。</param>
        ''' <summary>コンストラクタ。</summary>
        ''' <param name="target">パラメータ。</param>
        Public Sub New(target As Object)
            Me.mType = target?.GetType()
            Me.mTarget = target
            Me.Retrying = False
        End Sub

        ''' <summary>ローカル変数を消去します。</summary>
        Public Sub LocalVarClear() Implements IEnvironmentValue.LocalVarClear
            Me.mVariants.Clear()
        End Sub

        ''' <summary>ローカル変数を追加する。</summary>
        ''' <param name="name">変数名。</param>
        ''' <param name="value">変数値。</param>
        Public Sub AddVariant(name As String, value As Object) Implements IEnvironmentValue.AddVariant
            If Me.mVariants.ContainsKey(name) Then
                Me.mVariants(name) = value
            Else
                Me.mVariants.Add(name, value)
            End If
        End Sub

        ''' <summary>ローカル変数を削除する。</summary>
        ''' <param name="name">変数名。</param>
        Public Sub RemoveVariant(name As String) Implements IEnvironmentValue.RemoveVariant
            Me.mVariants.Remove(name)
        End Sub

        ''' <summary>指定した名称のプロパティから値を取得します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <returns>値。</returns>
        Public Function GetValue(name As String) As Object Implements IEnvironmentValue.GetValue
            If Me.mVariants.ContainsKey(name) Then
                ' ローカル変数に存在しているため、それを返す
                Return Me.mVariants(name)
            Else
                If Not Me.mPropDic.ContainsKey(name) Then
                    ' プロパティを取得して値を返す
                    Dim prop = Me.mType?.GetProperty(name)
                    If prop IsNot Nothing Then
                        Me.mPropDic.Add(name, prop)
                        Return prop.GetValue(Me.mTarget)

                    ElseIf Not Me.Retrying Then
                        Try
                            Me.Retrying = True
                            Dim tokens = SplitToken(name)
                            Dim ans = ParserAnalysis.Executes(tokens, Me)
                            Return ans.Contents

                        Catch ex As Exception
                            Logger.Value?.LogError("式を評価できません:{name}", name)
                        Finally
                            Me.Retrying = False
                        End Try
                    End If
                    Throw New DSqlAnalysisException($"指定したプロパティが定義されていません:{name}")
                Else
                    ' 保持しているプロパティ参照から値を返す
                    Return Me.mPropDic(name).GetValue(Me.mTarget)
                End If
            End If
        End Function

        ''' <summary>指定した名称のプロパティから値を取得します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <param name="index">インデックス。</param>
        ''' <returns>値。</returns>
        Public Function GetValue(name As String, index As Integer) As Object Implements IEnvironmentValue.GetValue
            Dim v = Me.GetValue(name)
            If TypeOf v Is IEnumerable Then
                Dim i = 0
                For Each e In CType(v, IEnumerable)
                    If i = index Then
                        Return e
                    End If
                    i += 1
                Next
                Throw New DSqlAnalysisException($"指定したインデックスが範囲外です:{index}")
            Else
                Throw New DSqlAnalysisException($"指定したプロパティは配列ではありません:{name}")
            End If
        End Function

        ''' <summary>指定した名称のメソッドを取得します。</summary>
        ''' <param name="name">メソッド名。</param>
        ''' <param name="value">値。</param>
        Public Shared Function GetMethod(name As String) As Func(Of Object(), IToken)
            Select Case name.ToLower()
                Case "format"
                    Return Function(args As Object())
                               If args.Length = 0 Then
                                   Throw New DSqlAnalysisException("format関数に引数を設定してください。")
                               ElseIf args.Length = 1 Then
                                   Return New StringToken(args(0).ToString(), """"c)
                               Else
                                   Return New StringToken(String.Format(args(0).ToString(), args.Skip(1).ToArray()), """"c)
                               End If
                           End Function
            End Select
            Throw New DSqlAnalysisException("有効な関数を指定していません。")
        End Function

        ''' <summary>指定した名称が定義されているかを取得します。</summary>
        ''' <param name="name">名称。</param>
        ''' <returns>真値の場合、定義されている。</returns>
        Public Function IsDefainedName(name As String) As Boolean Implements IEnvironmentValue.IsDefainedName
            If Me.mPropDic.ContainsKey(name) Then
                Return True
            End If

            If Me.mVariants.ContainsKey(name) Then
                Return True
            End If

            Return False
        End Function

        ''' <summary>環境値をコピーします。</summary>
        ''' <returns>コピーされた環境値。</returns>
        Public Function Clone() As IEnvironmentValue Implements IEnvironmentValue.Clone
            Return New EnvironmentObjectValue(Me)
        End Function

    End Class

End Namespace
