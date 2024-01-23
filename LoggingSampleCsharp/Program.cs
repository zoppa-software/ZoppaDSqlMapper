using Microsoft.Extensions.Logging;
using System.Runtime.InteropServices;
using ZoppaLoggingExtensions;

using var loggerFactory = LoggerFactory.Create(builder => {
    builder.AddZoppaLogging(
        (config) => config.MinimumLogLevel = LogLevel.Trace
    );
    builder.SetMinimumLevel(LogLevel.Trace);
});

ILogger logger = loggerFactory.CreateLogger<Program>();

var cls = new TestCls(loggerFactory);
cls.Log();

using (logger.BeginScope("scope start {a}", 100)) {
    logger.ZLog<Program>().LogDebug(1, "Does this line get hit? {h} {b}", 100, 200);
    using (logger.ZLog<Program>().BeginScope("2")) {
        logger.ZLog<Program>().LogDebug(3, "Nothing to see here.");
        logger.ZLog<Program>().LogDebug(5, "Warning... that was odd.");
    }
    logger.ZLog<Program>().LogDebug(7, "Oops, there was an error.");
}
logger.ZLog<Program>().LogDebug(5, "== 120.");

logger.ZLog<Program>().LogDebug(new EventId(), null, "Does this line get hit?", null);

class TestCls
{
    private readonly ILoggerFactory loggerFactory;

    public TestCls(ILoggerFactory loggerFactory)
    {
        this.loggerFactory = loggerFactory;
    }

    public void Log()
    {
        ILogger logger = loggerFactory.CreateLogger<TestCls>();
        logger.ZLog(this).LogDebug(1, "hit? {c:0.0} {d}", 100, 200);
        using (logger.ZLog(this).BeginScope("2")) {
            logger.ZLog(this).LogDebug(3, "Nothing to see here.");
            logger.ZLog(this).LogDebug(5, "Warning... that was odd.");
        }
        logger.ZLog(this).LogDebug(7, "Oops, there was an error.");
    }
}