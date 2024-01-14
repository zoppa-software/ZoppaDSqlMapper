Option Strict On
Option Explicit On

Namespace TokenCollection

    ''' <summary>トークン階層リンク。</summary>
    Friend NotInheritable Class TokenHierarchy

        ' 対象トークン
        Private mToken As TokenPosition

        ' 子要素リスト
        Private mChildren As List(Of TokenHierarchy)

        ''' <summary>対象トークンを取得します。</summary>
        Public ReadOnly Property TargetToken As TokenPosition
            Get
                Return Me.mToken
            End Get
        End Property

        ''' <summary>子要素リストを取得します。</summary>
        Public ReadOnly Property Children As List(Of TokenHierarchy)
            Get
                Return Me.mChildren
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="tkn">対象トークン。</param>
        Public Sub New(tkn As TokenPosition)
            Me.mToken = tkn
        End Sub

        ''' <summary>子要素を追加します。</summary>
        ''' <param name="token">対象トークン。</param>
        ''' <returns>追加した階層リンク。</returns>
        Friend Function AddChild(token As TokenPosition) As TokenHierarchy
            If Me.mChildren Is Nothing Then
                Me.mChildren = New List(Of TokenHierarchy)()
            End If
            Dim cnode = New TokenHierarchy(token)
            Me.mChildren.Add(cnode)
            Return cnode
        End Function

    End Class

End Namespace
