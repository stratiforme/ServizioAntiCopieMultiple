using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Hosting.WindowsServices;
using Microsoft.Extensions.Logging.EventLog;
using Serilog;
using Serilog.Events;
using ServizioAntiCopieMultiple;

string logsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple", "logs");
Directory.CreateDirectory(logsDir);
string logFilePath = Path.Combine(logsDir, "service-.log"); // Serilog rolling file pattern

Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Information()
    .Enrich.FromLogContext()
    .WriteTo.File(logFilePath, rollingInterval: RollingInterval.Day, retainedFileCountLimit: 14)
    .CreateLogger();

try
{
    IHost host = Host.CreateDefaultBuilder(args)
        .UseSerilog()
        .UseWindowsService(options =>
        {
            options.ServiceName = "ServizioAntiCopieMultiple";
        })
        .ConfigureLogging((context, logging) =>
        {
            logging.ClearProviders();
            logging.AddEventLog(new EventLogSettings
            {
                LogName = "Application",
                SourceName = "ServizioAntiCopieMultiple"
            });
            logging.AddSerilog();
        })
        .ConfigureServices(services =>
        {
            services.AddHostedService<PrintMonitorWorker>();
        })
        .Build();

    host.Run();
}
finally
{
    Log.CloseAndFlush();
}
