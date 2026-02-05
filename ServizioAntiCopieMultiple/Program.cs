using System;
using System.IO;
using System.Runtime.Versioning;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Hosting.WindowsServices;
using Microsoft.Extensions.Logging.EventLog;
using Serilog;
using Serilog.Events;
using ServizioAntiCopieMultiple;
using Microsoft.Extensions.Configuration;

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

// Determine minimum log level: default depends on run mode, but allow override via env var SACM_LOG_LEVEL
LogEventLevel defaultLevel = runAsService ? LogEventLevel.Information : LogEventLevel.Debug;
var envLevel = Environment.GetEnvironmentVariable("SACM_LOG_LEVEL");
if (!string.IsNullOrWhiteSpace(envLevel))
{
    if (Enum.TryParse<LogEventLevel>(envLevel, true, out var parsed))
    {
        defaultLevel = parsed;
    }
    else
    {
        // allow short names like "warn" or "warning"
        if (envLevel.Equals("warn", StringComparison.OrdinalIgnoreCase)) defaultLevel = LogEventLevel.Warning;
        else if (envLevel.Equals("error", StringComparison.OrdinalIgnoreCase)) defaultLevel = LogEventLevel.Error;
        else if (envLevel.Equals("information", StringComparison.OrdinalIgnoreCase) || envLevel.Equals("info", StringComparison.OrdinalIgnoreCase)) defaultLevel = LogEventLevel.Information;
        else if (envLevel.Equals("debug", StringComparison.OrdinalIgnoreCase)) defaultLevel = LogEventLevel.Debug;
        else if (envLevel.Equals("verbose", StringComparison.OrdinalIgnoreCase)) defaultLevel = LogEventLevel.Verbose;
    }
}

// Console sink can be forced via SACM_ENABLE_CONSOLE=true (useful when running interactively but starting without --console)
bool forceConsole = false;
var envConsole = Environment.GetEnvironmentVariable("SACM_ENABLE_CONSOLE");
if (!string.IsNullOrWhiteSpace(envConsole) && bool.TryParse(envConsole, out var cval)) forceConsole = cval;

// Always ensure file logging is enabled for diagnostics
bool ensureFileLogging = true;
var envFileLogging = Environment.GetEnvironmentVariable("SACM_DISABLE_FILE_LOGGING");
if (!string.IsNullOrWhiteSpace(envFileLogging) && bool.TryParse(envFileLogging, out var fval)) ensureFileLogging = !fval;

// Configure Serilog
var loggerConfig = new LoggerConfiguration()
    .MinimumLevel.Is(defaultLevel)
    .Enrich.FromLogContext();

// File logging is always enabled (for diagnostics), unless explicitly disabled
if (ensureFileLogging)
{
    loggerConfig = loggerConfig.WriteTo.File(logFilePath, rollingInterval: RollingInterval.Day, retainedFileCountLimit: 14);
}

// Console logging based on interactive mode or force flag
if (!runAsService || forceConsole)
{
    // add console sink for interactive debugging
    loggerConfig = loggerConfig.WriteTo.Console(restrictedToMinimumLevel: defaultLevel);
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
        .ConfigureAppConfiguration((ctx, cfg) =>
        {
            // allow appsettings.json to configure PrintMonitor
            cfg.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            // Also allow a global config placed in ProgramData by the installer / gestionesacm UI
            try
            {
                var commonPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple", "config.json");
                cfg.AddJsonFile(commonPath, optional: true, reloadOnChange: true);
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Could not add common ProgramData config.json to configuration sources");
            }
        })
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
