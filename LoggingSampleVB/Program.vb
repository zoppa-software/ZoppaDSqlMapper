Option Strict On
Option Explicit On

Imports System
Imports Microsoft.Extensions.DependencyInjection
Imports Microsoft.Extensions.Hosting
Imports Microsoft.Extensions.Logging
Imports ZoppaLoggingExtensions

Class Program

    Public Shared Sub Main(args As String())
        Task.Run(
            Async Function()
                Await MainAsync(args)
            End Function
        ).Wait()
    End Sub

    Private Shared Async Function MainAsync(args As String()) As Task
        Dim builder = Microsoft.Extensions.Hosting.Host.CreateApplicationBuilder(args)

        builder.Logging.ClearProviders()
        builder.Logging.AddZoppaLogging(
            Sub(config)
                config.MinimumLogLevel = LogLevel.Trace
            End Sub
        )
        builder.Logging.SetMinimumLevel(LogLevel.Trace)

        Using host = builder.Build()
            Dim loggerFactory = host.Services.GetRequiredService(Of ILoggerFactory)()

            Dim logger = loggerFactory.CreateLogger("Test")

            logger.ZLog(Of Program).LogDebug(1, "Does this line get hit?")
            logger.ZLog(Of Program).LogInformation(3, "Nothing to see here.")
            logger.ZLog(Of Program).LogWarning(5, "Warning... that was odd.")
            logger.ZLog(Of Program).LogError(7, "Oops, there was an error.")
            logger.ZLog(Of Program).LogTrace(5, "== 120.")

            Await host.RunAsync()
        End Using
    End Function

End Class
