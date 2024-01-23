Option Strict On
Option Explicit On

Imports System.IO
Imports System.IO.Compression
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Threading
Imports Microsoft.Extensions.Logging

''' <summary>ローカルログ出力クラスです。</summary>
NotInheritable Class LocalLogger

    ' デフォルトロガー
    Private Shared _defaultLogger As LocalLogger = Nothing

    ' 設定情報
    Private _config As ZoppaLoggingConfiguration

    ' 対象ファイル
    Private _logFile As FileInfo

    ' エンコード
    Private _encode As System.Text.Encoding

    ' 書込みバッファ
    Private ReadOnly _queue As New Queue(Of ILogData)()

    ' エラー書込みバッファ
    Private ReadOnly _errQueue As New Queue(Of ILogData)()

    ' 前回書込み完了日時
    Private _prevWriteDate As Date = Date.Now

    ' 書込み中フラグ
    Private _writing As Boolean

    ' コンソール出力ストリーム
    Private ReadOnly _outWriter As TextWriter = System.Console.Out

    ' コンソールエラー出力ストリーム
    Private ReadOnly _errorWriter As TextWriter = System.Console.Error

    ''' <summary>書き込み中状態を取得します。</summary>
    ''' <returns>書き込み中状態。</returns>
    Public ReadOnly Property IsWriting() As Boolean
        Get
            SyncLock Me
                Return (Me._queue.Count + Me._errQueue.Count > 0)
            End SyncLock
        End Get
    End Property

    ''' <summary>日付の変更でログを切り替えるならば真を返します。</summary>
    ''' <returns>切り替えるならば真。</returns>
    Private ReadOnly Property ChangeOfDate() As Boolean
        Get
            Return Me._config.SwitchByDay AndAlso
                    Me._prevWriteDate.Date < Date.Now.Date
        End Get
    End Property

    ''' <summary>内部ロガーを取得します。</summary>
    ''' <param name="categoryName">カテゴリ名。</param>
    ''' <param name="config">設定情報。</param>
    ''' <returns>内部ロガー。</returns>
    Friend Shared Function Create(categoryName As String, config As Func(Of ZoppaLoggingConfiguration)) As LocalLogger
        Dim res As New LocalLogger With {
            ._config = DirectCast(config().Clone(), ZoppaLoggingConfiguration)
        }

        Dim logPath As String = ""
        If Not res._config.LogFileByCategory.TryGetValue(categoryName, logPath) Then
            logPath = res._config.DefaultLogFile
            If _defaultLogger IsNot Nothing Then
                Return _defaultLogger
            End If
            _defaultLogger = res
        End If

        With res
            ._logFile = New FileInfo(logPath)
            If Not ._logFile.Directory.Exists Then
                ._logFile.Directory.Create()
            End If

            ._encode = System.Text.Encoding.GetEncoding(._config.EncodeName)
        End With
        Return res
    End Function

    ''' <summary>ログをバッファに溜める。</summary>
    ''' <param name="message">出力するログ。</param>
    Sub Stock(message As ILogData)
        ' 書き出す情報をため込む
        Dim cnt As Integer
        SyncLock Me
            Me._queue.Enqueue(message)
            cnt = Me._queue.Count
        End SyncLock

        ' キューにログが溜まっていたら少々待機
        Me.WaitFlushed(cnt, Me._config.CacheLimit)
    End Sub

    ''' <summary>ログをファイルに出力します。</summary>
    Sub Write(mutex As System.Threading.Mutex)
        ' 別スレッドでファイルに出力
        Dim running As Boolean = False
        SyncLock Me
            If Not Me._writing Then
                Me._writing = True
                running = True
            End If
        End SyncLock
        If running Then
            Task.Run(
                Sub()
                    Try
                        mutex.WaitOne()
                        Me.ThreadWrite()
                    Finally
                        mutex.ReleaseMutex()
                    End Try
                End Sub
            )
        End If
    End Sub

    ''' <summary>キューに溜まっているログを出力します。</summary>
    ''' <param name="cnt">キューのログ数。</param>
    ''' <param name="limit">払い出しリミット。</param>
    ''' <param name="loopCount">待機ループ回数。</param>
    ''' <param name="interval">待機ループインターバル。</param>
    Private Sub WaitFlushed(cnt As Integer,
                            limit As Integer,
                            Optional loopCount As Integer = 10,
                            Optional interval As Integer = 100)
        If cnt > limit Then
            For i As Integer = 0 To loopCount - 1
                Thread.Sleep(interval)

                SyncLock Me
                    cnt = Me._queue.Count
                End SyncLock
                If cnt < limit Then Exit For
            Next
        End If
    End Sub

    ''' <summary>ログをファイルに出力する。</summary>
    Private Sub ThreadWrite()
        Me.ArchiveHistories()

        Try
            Me._logFile.Refresh()
            Using sw As New StreamWriter(Me._logFile.FullName, True, Me._encode)
                Dim writed As Boolean
                Do
                    ' キュー内の文字列を取得
                    '
                    ' 2. キューにログ情報がある
                    '    対象ログレベル以上のログレベルを出力する場合、出力する
                    ' 3. キューにログ情報が空の場合はループを抜けてファイルストリームを閉じる
                    writed = False
                    Dim ln As ILogData = Nothing
                    Dim outd As Boolean = False
                    SyncLock Me
                        If Me._errQueue.Count > 0 Then                  ' 1
                            ln = Me._errQueue.Dequeue()
                            outd = True
                        ElseIf Me._queue.Count > 0 Then
                            ln = Me._queue.Dequeue()                    ' 2
                            outd = True
                        Else
                            Exit Do                                     ' 3
                        End If
                    End SyncLock

                    ' ファイルに書き出す
                    If ln IsNot Nothing Then
                        Try
                            If outd Then
                                Dim msg = ln.GetFormatMessage()
                                sw.WriteLine(msg)

                                Dim wr = If(ln.LogLv >= LogLevel.[Error], Me._errorWriter, Me._outWriter)
                                wr.WriteLine(msg)
                            End If
                        Catch ex As Exception
                            Me._errQueue.Enqueue(ln)
                        End Try
                        writed = True
                    End If

                    ' 出力した結果、ログファイルが最大サイズを超える場合、ループを抜けてストリームを閉じる
                    Me._logFile.Refresh()
                    If Me._logFile.Length > Me._config.MaxLogSize OrElse Me.ChangeOfDate Then
                        Exit Do
                    End If
                Loop While writed
                sw.Flush()
            End Using

            ' 上のループを抜けたとき実行中フラグを落とす
            SyncLock Me
                Me._writing = False
            End SyncLock

            Threading.Thread.Sleep(10)

        Catch ex As Exception
            SyncLock Me
                Me._writing = False
            End SyncLock
        Finally
            Me._prevWriteDate = Date.Now
        End Try
    End Sub

    ''' <summary>ログファイルが最大サイズを超えているか、日付が変わったかでログファイルを圧縮します。</summary>
    Private Sub ArchiveHistories()
        Me._logFile.Refresh()

        If Me._logFile.Exists AndAlso
                (Me._logFile.Length > Me._config.MaxLogSize OrElse Me.ChangeOfDate) Then
            Try
                Me._prevWriteDate = Date.Now

                ' ファイル名の要素を分割
                Dim ext = Path.GetExtension(Me._logFile.Name)
                Dim nm = Me._logFile.Name.Substring(0, Me._logFile.Name.Length - ext.Length)
                Dim tn = Date.Now.ToString("yyyyMMddHHmmssfff")

                Dim zipPath = New IO.FileInfo($"{Me._logFile.Directory.FullName}\{nm}_{tn}\{nm}{ext}")
                Try
                    ' 圧縮するフォルダを作成
                    If Not zipPath.Exists Then
                        zipPath.Directory.Create()
                    End If

                    ' ログファイルを圧縮
                    '
                    ' 1. 圧縮フォルダにログファイル移動、移動出来たら圧縮
                    ' 2. 現在のログファイルを圧縮
                    If Me.RetryableMove(zipPath) Then                                           ' 1
                        Dim compressFile = $"{zipPath.Directory.FullName}.zip"
                        ZipFile.CreateFromDirectory(zipPath.Directory.FullName, compressFile)   ' 2
                    End If

                Catch ex As Exception
                    Throw
                Finally
                    Directory.Delete($"{zipPath.Directory.FullName}", True)
                End Try

                ' 過去ファイルを整理
                Dim oldfiles = Directory.GetFiles(Me._logFile.Directory.FullName, $"{nm}*.zip").ToList()
                If oldfiles.Count > Me._config.LogGeneration Then
                    Me.ArchiveOldFiles(oldfiles)
                End If

            Catch ex As Exception
                SyncLock Me
                    Me._writing = False
                End SyncLock
            End Try
        End If
    End Sub

    ''' <summary>ログファイルを圧縮するフォルダへ移動する。</summary>
    ''' <param name="zipPath">移動先ファイルパス。</param>
    ''' <param name="retryCount">リトライ回数。</param>
    ''' <param name="retryInterval">リトライインターバル。</param>
    ''' <returns>移動に成功した場合はTrue、失敗した場合はFalse。</returns>
    Private Function RetryableMove(zipPath As FileInfo,
                                   Optional retryCount As Integer = 5,
                                   Optional retryInterval As Integer = 100) As Boolean
        Dim exx As Exception = Nothing

        For i As Integer = 0 To retryCount - 1
            Try
                File.Move(Me._logFile.FullName, zipPath.FullName)
                Return True
            Catch ex As Exception
                exx = ex
                Thread.Sleep(retryInterval)
            End Try
        Next

        Throw exx
    End Function

    ''' <summary>過去ファイルを整理する。</summary>
    ''' <param name="oldFiles">過去ファイルリスト。</param>
    Private Sub ArchiveOldFiles(oldFiles As List(Of String))
        Task.Run(
            Sub()
                ' 削除順にファイルをソート
                oldFiles.Sort()

                ' キャンセルされていなければ削除
                Try
                    Do While oldFiles.Count > Me._config.LogGeneration
                        If File.Exists(oldFiles.First()) Then
                            File.Delete(oldFiles.First())
                            oldFiles.RemoveAt(0)
                        End If
                    Loop
                Catch ex As Exception

                End Try
            End Sub
        )
    End Sub

    ''' <summary>ログ出力終了を待機します。</summary>
    Public Sub WaitFinish()
        For i As Integer = 0 To 5 * 60  ' 事情があって書き込めないとき無限ループするためループ回数制限する
            If Me.IsWriting Then
                Me.FlushWrite()
                Threading.Thread.Sleep(1000)
            Else
                Exit For
            End If
        Next
    End Sub

    ''' <summary>出力スレッドが停止中ならば実行します。</summary>
    Private Sub FlushWrite()
        Try
            ' 出力スレッドが停止中ならばスレッド開始
            Dim running = False
            SyncLock Me
                If Not Me._writing Then
                    Me._writing = True
                    running = True
                End If
            End SyncLock
            If running Then
                Task.Run(Sub() Me.ThreadWrite())
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Interface ILogData

        ReadOnly Property LogLv As LogLevel

        Function GetFormatMessage() As String

    End Interface

End Class
