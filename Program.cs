using System;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

var host = Host.CreateDefaultBuilder(args)
    .UseWindowsService() // se è un Windows Service
    .ConfigureServices((context, services) =>
    {
        // servizi...
    })
    .Build();

var logger = host.Services.GetRequiredService<ILogger<Program>>();
var lifetime = host.Services.GetRequiredService<IHostApplicationLifetime>();

lifetime.ApplicationStopping.Register(() =>
{
    logger.LogWarning("Host.ApplicationStopping invoked - shutdown requested.");
});

lifetime.ApplicationStopped.Register(() =>
{
    logger.LogWarning("Host.ApplicationStopped invoked - shutdown complete.");
});

AppDomain.CurrentDomain.UnhandledException += (s, e) =>
{
    logger.LogCritical(e.ExceptionObject as Exception, "Unhandled exception (AppDomain).");
};

TaskScheduler.UnobservedTaskException += (s, e) =>
{
    logger.LogCritical(e.Exception, "UnobservedTaskException.");
    e.SetObserved();
};

await host.RunAsync();