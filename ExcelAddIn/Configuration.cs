using Serilog.Core;
using System.Reflection;

namespace HcExcelAddIn;

internal static class Configuration
{
    internal static IServiceProvider ConfigureServices(ExcelApplication excelApplication)
    {
        var logger = CreateLogger();
        var services = new ServiceCollection();

        services.AddSingleton(excelApplication);
        services.AddSingleton<ILogger>(logger);
        services.AddSingleton<IRibbonController, RibbonController>();
        services.AddSingleton<MainView>(provider => new MainView());

        return services.BuildServiceProvider();
    }

    private static Logger CreateLogger()
    {
        var logPath = GetLogFilePath();

        return new LoggerConfiguration()
            .MinimumLevel.Debug() // Change to Information or Warning in production
            .Enrich.FromLogContext()
            .WriteTo.File(
                path: logPath,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj} {Properties}{NewLine}{Exception}",
                rollingInterval: RollingInterval.Day,
                restrictedToMinimumLevel: LogEventLevel.Debug,
                retainedFileCountLimit: 7
            )
            .CreateLogger();
    }

    private static string GetLogFilePath()
    {
        var basePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        var safePath = basePath ?? Path.Combine(Environment.CurrentDirectory, "logs");
        Directory.CreateDirectory(safePath);

        return Path.Combine(safePath, "Addin-.log");
    }
}
