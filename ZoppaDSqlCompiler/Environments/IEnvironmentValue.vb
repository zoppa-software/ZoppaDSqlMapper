Option Strict On
Option Explicit On

Namespace Environments

    ''' <summary>環境値情報。</summary>
    Public Interface IEnvironmentValue

        ''' <summary>ローカル変数を消去します。</summary>
        Sub LocalVarClear()

        ''' <summary>ローカル変数を追加する。</summary>
        ''' <param name="name">変数名。</param>
        ''' <param name="value">変数値。</param>
        Sub AddVariant(name As String, value As Object)

        ''' <summary>指定した名称のプロパティから値を取得します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <returns>値。</returns>
        Function GetValue(name As String) As Object

    End Interface

End Namespace