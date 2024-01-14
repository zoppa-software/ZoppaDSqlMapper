Option Strict On
Option Explicit On

Namespace TokenCollection

    ''' <summary>入力トークンストリーム。</summary>
    Friend NotInheritable Class TokenStream

        ''' <summary>シーク位置ポインタ。</summary>
        Private mPointer As Integer

        ''' <summary>入力トークン。</summary>
        Private mTokens As TokenPosition()

        ''' <summary>読み込みの終了していない文字があれば真を返す。</summary>
        Public ReadOnly Property HasNext As Boolean
            Get
                Return (Me.mPointer < If(Me.mTokens?.Length, 0))
            End Get
        End Property

        ''' <summary>カレント文字を返す。</summary>
        Public ReadOnly Property Current As TokenPosition
            Get
                Return If(Me.mPointer < If(Me.mTokens?.Length, 0), Me.mTokens(Me.mPointer), Nothing)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputtkn">入力トークン。</param>
        Public Sub New(inputtkn As IEnumerable(Of TokenPosition))
            Me.mPointer = 0
            Me.mTokens = inputtkn.ToArray()
        End Sub

        ''' <summary>カレント位置を移動させる。</summary>
        ''' <param name="moveAmount">移動量。</param>
        Public Sub Move(moveAmount As Integer)
            Me.mPointer += moveAmount
        End Sub

    End Class

End Namespace

