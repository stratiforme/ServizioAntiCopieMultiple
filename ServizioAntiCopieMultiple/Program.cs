using System.Runtime.Versioning;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Hosting.WindowsServices;
using Microsoft.Extensions.Logging.EventLog;
using Serilog;
using Serilog.Events;
using ServizioAntiCopieMultiple;

[assembly: SupportedOSPlatform("windows")]

bool runAsService = !(Environment.UserInteractive || args != null && args.Length > 0 && args[0] == "--console");

string logsDir;
try
{
    logsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple", "logs");
    Directory.CreateDirectory(logsDir);
}
catch (Exception)
{
    // Fallback to temp path if creating under ProgramData fails (insufficient permissions)
    logsDir = Path.Combine(Path.GetTempPath(), "ServizioAntiCopieMultiple", "logs");
    Directory.CreateDirectory(logsDir);
}

string logFilePath = Path.Combine(logsDir, "service-.log"); // Serilog rolling file pattern

var loggerConfig = new LoggerConfiguration()
    .MinimumLevel.Information()
    .Enrich.FromLogContext()
    .WriteTo.File(logFilePath, rollingInterval: RollingInterval.Day, retainedFileCountLimit: 14);

if (!runAsService)
{
    // add console sink for interactive debugging
    loggerConfig = loggerConfig.WriteTo.Console();
}

Log.Logger = loggerConfig.CreateLogger();

// Global exception handlers to ensure we log fatal crashes
AppDomain.CurrentDomain.UnhandledException += (s, e) =>
{
    Log.Fatal(e.ExceptionObject as Exception, "Unhandled exception (AppDomain)");
    Log.CloseAndFlush();
};

TaskScheduler.UnobservedTaskException += (s, e) =>
{
    Log.Error(e.Exception, "Unobserved task exception");
};

try
{
    var builder = Host.CreateDefaultBuilder(args)
        .UseSerilog()
        .ConfigureLogging((context, logging) =>
        {
            logging.ClearProviders();
            try
            {
                logging.AddEventLog(new EventLogSettings
                {
                    LogName = "Application",
                    SourceName = "ServizioAntiCopieMultiple"
                });
            }
            catch (Exception ex)
            {
                // If EventLog registration fails (lack of permissions), continue without it
                Log.Warning(ex, "Impossibile registrare EventLog provider; continuerà senza EventLog logging.");
            }
        })
        .ConfigureServices(services =>
        {
            services.AddHostedService<PrintMonitorWorker>();
        });

    if (runAsService)
    {
        builder = builder.UseWindowsService(options => { options.ServiceName = "ServizioAntiCopieMultiple"; });
    }

    using IHost host = builder.Build();

    if (runAsService)
    {
        host.Run();
    }
    else
    {
        // Run as console app for easier testing and diagnostics
        await host.RunAsync();
    }
}
finally
{
    Log.CloseAndFlush();
}
