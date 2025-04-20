using System.Reflection;

namespace HcExcelAddIn;

internal static class Configuration
{
    internal static ILogger ConfigureLogger()
    {
        var folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? $"logs";
        var logFilePath = Path.Combine(folder, "Addin-.log");

        var logger = new LoggerConfiguration()
            .MinimumLevel.Debug() // Change to Information or Warning in production
            .Enrich.FromLogContext()
            .WriteTo.File(
                path: logFilePath,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj} {Properties}{NewLine}{Exception}",
                rollingInterval: RollingInterval.Day,
                restrictedToMinimumLevel: LogEventLevel.Debug,      // Change to Information or Warning in production
                retainedFileCountLimit: 7 // Keep logs for last X days
            )
            .CreateLogger();

        logger.Debug("Logger configured.");

        return logger;
    }
}
