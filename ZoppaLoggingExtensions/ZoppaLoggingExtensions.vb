Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports Microsoft.Extensions.DependencyInjection
Imports Microsoft.Extensions.DependencyInjection.Extensions
Imports Microsoft.Extensions.Logging
Imports Microsoft.Extensions.Logging.Configuration

''' <summary>ログ出力拡張クラスです。</summary>
Public Module ZoppaLogProcessExtensions

    ' ログ出力フォーマッタ
    Private ReadOnly _messageFormatter As Func(Of ZoppaFormattedLogValues, Exception, String) = AddressOf MessageFormatter

    ''' <summary>サービスの登録を行います。</summary>
    ''' <param name="builder">ログビルダー。</param>
    ''' <returns>ログビルダー。</returns>
    <Extension()>
    Public Function AddZoppaLogging(builder As ILoggingBuilder) As ILoggingBuilder
        builder.AddConfiguration()

        builder.Services.TryAddEnumerable(
            ServiceDescriptor.Singleton(Of ILoggerProvider, ZoppaLoggingProvider)()
        )

        LoggerProviderOptions.RegisterProviderOptions(Of ZoppaLoggingConfiguration, ZoppaLoggingProvider)(builder.Services)

        Return builder
    End Function

    ''' <summary>サービスの登録を行います。</summary>
    ''' <param name="builder">ログビルダー。</param>
    ''' <param name="configure">ログ設定。</param>
    ''' <returns>ログビルダー。</returns>
    <Extension()>
    Public Function AddZoppaLogging(builder As ILoggingBuilder,
                                    configure As Action(Of ZoppaLoggingConfiguration)) As ILoggingBuilder
        builder.AddZoppaLogging()
        builder.Services.Configure(configure)

        Return builder
    End Function

#Region "debug"

    ''' <summary>フォーマットしてデバッグログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogDebug(0, exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogDebug(logger As LogWrapper, eventId As EventId, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Debug, eventId, exception, message, args)
    End Sub

    ''' <summary>フォーマットしてデバッグログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogDebug(0, "Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogDebug(logger As LogWrapper, eventId As EventId, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Debug, eventId, Nothing, message, args)
    End Sub

    ''' <summary>フォーマットしてデバッグログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogDebug(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogDebug(logger As LogWrapper, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Debug, 0, exception, message, args)
    End Sub

    ''' <summary>フォーマットしてデバッグログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <example>logger.LogDebug(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogDebug(logger As LogWrapper, exception As Exception)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Debug, 0, exception, "", Array.Empty(Of Object)())
    End Sub

    ''' <summary>フォーマットしてデバッグログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogDebug("Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogDebug(logger As LogWrapper, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Debug, 0, Nothing, message, args)
    End Sub

#End Region

#Region "trace"

    ''' <summary>フォーマットしてトレースログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogTrace(0, exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogTrace(logger As LogWrapper, eventId As EventId, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Trace, eventId, exception, message, args)
    End Sub

    ''' <summary>フォーマットしてトレースログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogTrace(0, "Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogTrace(logger As LogWrapper, eventId As EventId, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Trace, eventId, Nothing, message, args)
    End Sub

    ''' <summary>フォーマットしてトレースログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogTrace(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogTrace(logger As LogWrapper, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Trace, 0, exception, message, args)
    End Sub

    ''' <summary>フォーマットしてトレースログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <example>logger.LogTrace(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogTrace(logger As LogWrapper, exception As Exception)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Trace, 0, exception, "", Array.Empty(Of Object)())
    End Sub

    ''' <summary>フォーマットしてトレースログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogTrace("Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogTrace(logger As LogWrapper, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Trace, 0, Nothing, message, args)
    End Sub

#End Region

#Region "information"

    ''' <summary>フォーマットして通常ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogInformation(0, exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogInformation(logger As LogWrapper, eventId As EventId, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Information, eventId, exception, message, args)
    End Sub

    ''' <summary>フォーマットして通常ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogInformation(0, "Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogInformation(logger As LogWrapper, eventId As EventId, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Information, eventId, Nothing, message, args)
    End Sub

    ''' <summary>フォーマットして通常ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogInformation(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogInformation(logger As LogWrapper, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Information, 0, exception, message, args)
    End Sub

    ''' <summary>フォーマットして通常ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <example>logger.LogInformation(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogInformation(logger As LogWrapper, exception As Exception)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Information, 0, exception, "", Array.Empty(Of Object)())
    End Sub

    ''' <summary>フォーマットして通常ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogInformation("Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogInformation(logger As LogWrapper, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Information, 0, Nothing, message, args)
    End Sub

#End Region

#Region "warning"

    ''' <summary>フォーマットして警告ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogWarning(0, exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogWarning(logger As LogWrapper, eventId As EventId, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Warning, eventId, exception, message, args)
    End Sub

    ''' <summary>フォーマットして警告ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogWarning(0, "Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogWarning(logger As LogWrapper, eventId As EventId, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Warning, eventId, Nothing, message, args)
    End Sub

    ''' <summary>フォーマットして警告ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogWarning(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogWarning(logger As LogWrapper, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Warning, 0, exception, message, args)
    End Sub

    ''' <summary>フォーマットして警告ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <example>logger.LogWarning(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogWarning(logger As LogWrapper, exception As Exception)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Warning, 0, exception, "", Array.Empty(Of Object)())
    End Sub

    ''' <summary>フォーマットして警告ログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogWarning("Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogWarning(logger As LogWrapper, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Warning, 0, Nothing, message, args)
    End Sub

#End Region

#Region "error"

    ''' <summary>フォーマットしてエラーログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogError(0, exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogError(logger As LogWrapper, eventId As EventId, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Error, eventId, exception, message, args)
    End Sub

    ''' <summary>フォーマットしてエラーログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogError(0, "Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogError(logger As LogWrapper, eventId As EventId, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Error, eventId, Nothing, message, args)
    End Sub

    ''' <summary>フォーマットしてエラーログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogError(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogError(logger As LogWrapper, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Error, 0, exception, message, args)
    End Sub

    ''' <summary>フォーマットしてエラーログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <example>logger.LogError(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogError(logger As LogWrapper, exception As Exception)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Error, 0, exception, "", Array.Empty(Of Object)())
    End Sub

    ''' <summary>フォーマットしてエラーログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogError("Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogError(logger As LogWrapper, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Error, 0, Nothing, message, args)
    End Sub

#End Region

#Region "critical"

    ''' <summary>フォーマットしてCriticalなログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogCritical(0, exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogCritical(logger As LogWrapper, eventId As EventId, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Critical, eventId, exception, message, args)
    End Sub

    ''' <summary>フォーマットしてCriticalなログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogCritical(0, "Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogCritical(logger As LogWrapper, eventId As EventId, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Critical, eventId, Nothing, message, args)
    End Sub

    ''' <summary>フォーマットしてCriticalなログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogCritical(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogCritical(logger As LogWrapper, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Critical, 0, exception, message, args)
    End Sub

    ''' <summary>フォーマットしてCriticalなログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <example>logger.LogCritical(exception, "Error while processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogCritical(logger As LogWrapper, exception As Exception)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Critical, 0, exception, "", Array.Empty(Of Object)())
    End Sub

    ''' <summary>フォーマットしてCriticalなログを出力します、</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <example>logger.LogCritical("Processing request from {Address}", address)</example>
    <Extension()>
    Public Sub LogCritical(logger As LogWrapper, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, LogLevel.Critical, 0, Nothing, message, args)
    End Sub

#End Region

#Region "common"

    ''' <summary>ログラッパーを取得します。</summary>
    ''' <typeparam name="T">ログ出力するクラス。</typeparam>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。
    ''' <param name="caller">呼び出し元のオブジェクト。</param>
    ''' <param name="callMember">呼び出し元のメソッド。</param>
    ''' <param name="lineNumber">呼び出し行位置。</param>
    ''' <returns>ログラッパー。</returns>
    <Extension()>
    Public Function ZLog(Of T)(logger As ILogger, caller As T, <CallerMemberName> Optional callMember As String = "", <CallerLineNumber> Optional lineNumber As Integer = 0) As LogWrapper
        Return New LogWrapper(logger, GetType(T), callMember, lineNumber)
    End Function

    ''' <summary>ログラッパーを取得します。</summary>
    ''' <typeparam name="T">ログ出力するクラス。</typeparam>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。
    ''' <param name="callMember">呼び出し元のメソッド。</param>
    ''' <param name="lineNumber">呼び出し行位置。</param>
    ''' <returns>ログラッパー。</returns>
    <Extension()>
    Public Function ZLog(Of T)(logger As ILogger, <CallerMemberName> Optional callMember As String = "", <CallerLineNumber> Optional lineNumber As Integer = 0) As LogWrapper
        Return New LogWrapper(logger, GetType(T), callMember, lineNumber)
    End Function

    ''' <summary>フォーマット また、指定されたログレベルでログメッセージを書き込みます。</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="logLevel">書き込むログレベル。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    <Extension()>
    Private Sub Log(logger As LogWrapper, logLevel As LogLevel, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, logLevel, 0, Nothing, message, args)
    End Sub

    ''' <summary>フォーマット また、指定されたログレベルでログメッセージを書き込みます。</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="logLevel">書き込むログレベル。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    <Extension()>
    Public Sub Log(logger As LogWrapper, logLevel As LogLevel, eventId As EventId, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, logLevel, eventId, Nothing, message, args)
    End Sub

    ''' <summary>フォーマット また、指定されたログレベルでログメッセージを書き込みます。</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="logLevel">書き込むログレベル。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    <Extension()>
    Public Sub Log(logger As LogWrapper, logLevel As LogLevel, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, logLevel, 0, exception, message, args)
    End Sub

    ''' <summary>フォーマット また、指定されたログレベルでログメッセージを書き込みます。</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="logLevel">書き込むログレベル。</param>
    ''' <param name="exception">書き込む例外。</param>
    <Extension()>
    Public Sub Log(logger As LogWrapper, logLevel As LogLevel, exception As Exception)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, logLevel, 0, exception, "", Array.Empty(Of Object)())
    End Sub

    ''' <summary>フォーマット また、指定されたログレベルでログメッセージを書き込みます。</summary>
    ''' <param name="logger">この<see cref="ILogger"/>に出力します。</param>
    ''' <param name="logLevel">書き込むログレベル。</param>
    ''' <param name="eventId">ログに関連付けられているイベント ID。</param>
    ''' <param name="exception">書き込む例外。</param>
    ''' <param name="message">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    <Extension()>
    Public Sub Log(logger As LogWrapper, logLevel As LogLevel, eventId As EventId, exception As Exception, message As String, ParamArray args() As Object)
        Dim frame = New System.Diagnostics.StackTrace(1, True).GetFrame(0)
        Log(logger, frame, logLevel, eventId, exception, message, args)
    End Sub

    ' ログ出力機能
    Private Sub Log(wrap As LogWrapper, frame As StackFrame, logLevel As LogLevel, eventId As EventId, exception As Exception, message As String, ParamArray args() As Object)
        If wrap.Logger IsNot Nothing Then
            wrap.Logger.Log(logLevel, eventId, New ZoppaFormattedLogValues(message, wrap.CallerType, wrap.CallerMember, wrap.CallerLineNumber, args), exception, _messageFormatter)
        Else
            Throw New NullReferenceException("loggerの参照がNullです。")
        End If
    End Sub

    ''' <summary>メッセージをフォーマットして、スコープを作成します。</summary>
    ''' <param name="wrap">この<see cref="ILogger"/>のスコープを作成します。</param>
    ''' <param name="messageFormat">ログメッセージのフォーマット文字列。例は <c>"User {User} logged in from {Address}"</c></param>
    ''' <param name="args">書式設定する 0 個以上のオブジェクトを含むオブジェクト配列。</param>
    ''' <returns>破棄可能なスコープ オブジェクト。</returns>
    ''' <example>
    ''' using(logger.ZBeginScope("Processing request from {Address}", address))
    ''' {
    ''' }
    ''' </example>
    <Extension()>
    Public Function BeginScope(wrap As LogWrapper, messageFormat As String, ParamArray args() As Object) As IDisposable
        If wrap.Logger IsNot Nothing Then
            Return wrap.Logger.BeginScope(New ZoppaFormattedLogValues(messageFormat, wrap.CallerType, wrap.CallerMember, wrap.CallerLineNumber, args))
        Else
            Throw New NullReferenceException("logger")
        End If
    End Function

    Private Function MessageFormatter(state As ZoppaFormattedLogValues, ex As Exception) As String
        If ex Is Nothing Then
            Return state.LogMessage()
        Else
            Return state.LogMessage(ex)
        End If
    End Function

#End Region

    ''' <summary>ログラッパー。</summary>
    Public Structure LogWrapper

        ''' <summary>この<see cref="ILogger"/>に出力します。</summary>
        Public ReadOnly Logger As ILogger

        ''' <summary>呼び出し元のオブジェクト。</summary>
        Public ReadOnly CallerType As Type

        ''' <summary>呼び出し元のメソッド。</summary>
        Public ReadOnly CallerMember As String

        ''' <summary>呼び出し行位置。</summary>
        Public ReadOnly CallerLineNumber As Integer

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="log">ログ。</param>
        ''' <param name="ty">呼び出し元のオブジェクト。</param>
        ''' <param name="mem">呼び出し元のメソッド。</param>
        ''' <param name="ln">呼び出し行位置。</param>
        Public Sub New(log As ILogger, ty As Type, mem As String, ln As Integer)
            Me.Logger = log
            Me.CallerType = ty
            Me.CallerMember = mem
            Me.CallerLineNumber = ln
        End Sub

    End Structure

End Module