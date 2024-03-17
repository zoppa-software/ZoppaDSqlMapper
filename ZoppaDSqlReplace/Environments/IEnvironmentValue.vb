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

        ''' <summary>ローカル変数を削除する。</summary>
        ''' <param name="name">変数名。</param>
        Sub RemoveVariant(name As String)

        ''' <summary>指定した名称のプロパティから値を取得します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <returns>値。</returns>
        Function GetValue(name As String) As Object

        ''' <summary>指定した名称のプロパティから値を取得します。</summary>
        ''' <param name="name">プロパティ名。</param>
        ''' <param name="index">インデックス。</param>
        ''' <returns>値。</returns>
        Function GetValue(name As String, index As Integer) As Object

        ''' <summary>指定した名称が定義されているかを取得します。</summary>
        ''' <param name="name">名称。</param>
        ''' <returns>真値の場合、定義されている。</returns>
        Function IsDefainedName(name As String) As Boolean

        ''' <summary>環境値をコピーします。</summary>
        ''' <returns>コピーされた環境値。</returns>
        Function Clone() As IEnvironmentValue

    End Interface

End Namespace