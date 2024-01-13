Option Strict On
Option Explicit On

Imports System.Data
Imports System.Runtime.CompilerServices

''' <summary>SQL実行用パラメータ APIです。</summary>
Public Module ZoppaDSqlSetting

    ' タイムアウト設定
    Private mDefTimeOut As Integer = 30

    ''' <summary>デフォルトのタイムアウト値を設定、取得します。</summary>
    ''' <returns>タイムアウト値。</returns>
    Public Property DefaultTimeoutSecond As Integer
        Get
            Return mDefTimeOut
        End Get
        Set(value As Integer)
            If value >= -1 Then
                mDefTimeOut = value
            Else
                Throw New ZoppaDSqlException("タイムアウト値に不正な値を指定しています。")
            End If
        End Set
    End Property

    ''' <summary>デフォルトのSQLパラメータチェックを設定、取得します。</summary>
    ''' <returns>パラメータチェック式。</returns>
    Public Property DefaultSqlParameterCheck As Action(Of IDbDataParameter) = Nothing

    ''' <summary>デフォルトパラメータ接頭辞を設定、取得します。</summary>
    ''' <returns>パラメータ接頭辞。</returns>
    Public Property DefaultParameterPrefix As PrefixType = PrefixType.AtMark

    ''' <summary>タイムアウト値を設定します。</summary>
    ''' <param name="dbConnection">DBコネクション。</param>
    ''' <param name="timeout">タイムアウト値。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetTimeoutSecond(dbConnection As IDbConnection, timeout As Integer) As Settings
        Return New Settings(dbConnection, timeout, Nothing, DefaultParameterPrefix, CommandType.Text, DefaultSqlParameterCheck, New String() {})
    End Function

    ''' <summary>タイムアウト値を設定します。</summary>
    ''' <param name="prevSetting">設定済み情報。</param>
    ''' <param name="timeout">タイムアウト値。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetTimeoutSecond(prevSetting As Settings, timeout As Integer) As Settings
        With prevSetting
            Return New Settings(.DbConnection, timeout, .Transaction, .ParameterPrefix, .CommandType, .ParameterChecker, .PropertyNames)
        End With
    End Function

    ''' <summary>トランザクションを設定します。</summary>
    ''' <param name="dbConnection">DBコネクション。</param>
    ''' <param name="transaction">トランザクション。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetTransaction(dbConnection As IDbConnection, transaction As IDbTransaction) As Settings
        Return New Settings(dbConnection, DefaultTimeoutSecond, transaction, DefaultParameterPrefix, CommandType.Text, DefaultSqlParameterCheck, New String() {})
    End Function

    ''' <summary>トランザクションを設定します。</summary>
    ''' <param name="prevSetting">設定済み情報。</param>
    ''' <param name="transaction">トランザクション。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetTransaction(prevSetting As Settings, transaction As IDbTransaction) As Settings
        With prevSetting
            Return New Settings(.DbConnection, .TimeOutSecond, transaction, .ParameterPrefix, .CommandType, .ParameterChecker, .PropertyNames)
        End With
    End Function

    ''' <summary>使用するパラメータ接頭辞を指定します。</summary>
    ''' <param name="dbConnection">DBコネクション。</param>
    ''' <param name="prefix">パラメータ接頭辞。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetParameterPrepix(dbConnection As IDbConnection, prefix As PrefixType) As Settings
        Return New Settings(dbConnection, DefaultTimeoutSecond, Nothing, prefix, CommandType.Text, DefaultSqlParameterCheck, New String() {})
    End Function

    ''' <summary>使用するパラメータ接頭辞を指定します。</summary>
    ''' <param name="prevSetting">設定済み情報。</param>
    ''' <param name="prefix">パラメータ接頭辞。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetParameterPrepix(prevSetting As Settings, prefix As PrefixType) As Settings
        With prevSetting
            Return New Settings(.DbConnection, .TimeOutSecond, .Transaction, prefix, .CommandType, .ParameterChecker, .PropertyNames)
        End With
    End Function

    ''' <summary>コマンドタイプを指定します。</summary>
    ''' <param name="dbConnection">DBコネクション。</param>
    ''' <param name="cmdType">コマンドタイプ。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetCommandType(dbConnection As IDbConnection, cmdType As CommandType) As Settings
        Return New Settings(dbConnection, DefaultTimeoutSecond, Nothing, DefaultParameterPrefix, cmdType, DefaultSqlParameterCheck, New String() {})
    End Function

    ''' <summary>コマンドタイプを指定します。</summary>
    ''' <param name="prevSetting">設定済み情報。</param>
    ''' <param name="cmdType">コマンドタイプ。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetCommandType(prevSetting As Settings, cmdType As CommandType) As Settings
        With prevSetting
            Return New Settings(.DbConnection, .TimeOutSecond, .Transaction, .ParameterPrefix, cmdType, .ParameterChecker, .PropertyNames)
        End With
    End Function

    ''' <summary>SQLパラメータチェックを指定します。</summary>
    ''' <param name="dbConnection">DBコネクション。</param>
    ''' <param name="checker">SQLパラメータチェック。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetParameterChecker(dbConnection As IDbConnection, checker As Action(Of IDbDataParameter)) As Settings
        Return New Settings(dbConnection, DefaultTimeoutSecond, Nothing, DefaultParameterPrefix, CommandType.Text, checker, New String() {})
    End Function

    ''' <summary>SQLパラメータチェックを指定します。</summary>
    ''' <param name="prevSetting">設定済み情報。</param>
    ''' <param name="checker">SQLパラメータチェック。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetParameterChecker(prevSetting As Settings, checker As Action(Of IDbDataParameter)) As Settings
        With prevSetting
            Return New Settings(.DbConnection, .TimeOutSecond, .Transaction, .ParameterPrefix, .CommandType, checker, .PropertyNames)
        End With
    End Function

    ''' <summary>SQLパラメータを割り当てるプロパティ名を設定します。</summary>
    ''' <param name="dbConnection">DBコネクション。</param>
    ''' <param name="names">プロパティ名リスト。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetOrderName(dbConnection As IDbConnection, ParamArray names() As String) As Settings
        Return New Settings(dbConnection, DefaultTimeoutSecond, Nothing, DefaultParameterPrefix, CommandType.Text, DefaultSqlParameterCheck, names)
    End Function

    ''' <summary>SQLパラメータを割り当てるプロパティ名を設定します。</summary>
    ''' <param name="prevSetting">設定済み情報。</param>
    ''' <param name="names">プロパティ名リスト。</param>
    ''' <returns>設定値。</returns>
    <Extension()>
    Public Function SetOrderName(prevSetting As Settings, ParamArray names() As String) As Settings
        With prevSetting
            Return New Settings(.DbConnection, .TimeOutSecond, .Transaction, .ParameterPrefix, .CommandType, .ParameterChecker, names)
        End With
    End Function

    ''' <summary>SQL実行用パラメータ定義です。</summary>
    Public Structure Settings

        ''' <summary>DB接続情報です。</summary>
        Friend ReadOnly DbConnection As IDbConnection

        ' タイムアウト設定
        Private ReadOnly mTimeOut As Integer

        ' トランザクション設定
        Private ReadOnly mTran As IDbTransaction

        ' 変数接頭辞設定
        Private ReadOnly mVarPrefix As PrefixType

        ' コマンドタイプ
        Private ReadOnly mCmdType As CommandType

        ' SQLパラメータチェック
        Private ReadOnly mPrmCheck As Action(Of IDbDataParameter)

        ' プロパティ名リスト
        Private ReadOnly mPropNames As String()

        ''' <summary>タイムアウト設定を取得します。</summary>
        ''' <returns>タイムアウト時間（秒数）</returns>
        Public ReadOnly Property TimeOutSecond As Integer
            Get
                Return Me.mTimeOut
            End Get
        End Property

        ''' <summary>トランザクションを取得します。</summary>
        ''' <returns>トランザクションオブジェクト。</returns>
        Public ReadOnly Property Transaction As IDbTransaction
            Get
                Return Me.mTran
            End Get
        End Property

        ''' <summary>パラメータの接頭辞を取得します。</summary>
        ''' <returns>接頭辞の種類。</returns>
        Public ReadOnly Property ParameterPrefix As PrefixType
            Get
                Return Me.mVarPrefix
            End Get
        End Property

        ''' <summary>コマンドタイプを取得します。</summary>
        ''' <returns>コマンドタイプ。</returns>
        Public ReadOnly Property CommandType As CommandType
            Get
                Return Me.mCmdType
            End Get
        End Property

        ''' <summary>SQLパラメータチェックを取得します。</summary>
        ''' <returns>SQLパラメータチェック。</returns>
        Public ReadOnly Property ParameterChecker As Action(Of IDbDataParameter)
            Get
                Return Me.mPrmCheck
            End Get
        End Property

        ''' <summary>プロパティ名リストを取得します。</summary>
        ''' <returns>プロパティ名リスト。</returns>
        Public ReadOnly Property PropertyNames As String()
            Get
                Return Me.mPropNames
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="dbCon">DBコネクション。</param>
        Friend Sub New(dbCon As IDbConnection)
            Me.DbConnection = dbCon
            Me.mTimeOut = DefaultTimeoutSecond
            Me.mTran = Nothing
            Me.mVarPrefix = DefaultParameterPrefix
            Me.mCmdType = CommandType.Text
            Me.mPrmCheck = Nothing
            Me.mPropNames = New String() {}
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="dbCon">DBコネクション。</param>
        ''' <param name="timeOut">タイムアウト値。</param>
        ''' <param name="tran">トランザクション。</param>
        ''' <param name="varPrex">パラメータ接頭辞。</param>
        ''' <param name="cmdType">SQLコマンドのタイプ</param>
        ''' <param name="prmCheck">SQLパラメータチェック。</param>
        ''' <param name="propNames">プロパティ名リスト。</param>
        Friend Sub New(dbCon As IDbConnection,
                       timeOut As Integer,
                       tran As IDbTransaction,
                       varPrex As PrefixType,
                       cmdType As CommandType,
                       prmCheck As Action(Of IDbDataParameter),
                       propNames As String())
            Me.DbConnection = dbCon
            Me.mTimeOut = timeOut
            Me.mTran = tran
            Me.mVarPrefix = varPrex
            Me.mCmdType = cmdType
            Me.mPrmCheck = prmCheck
            Me.mPropNames = propNames
        End Sub

    End Structure

End Module

