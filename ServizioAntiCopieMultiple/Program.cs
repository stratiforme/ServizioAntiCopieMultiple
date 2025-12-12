using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Hosting.WindowsServices;
using ServizioAntiCopieMultiple;

IHost host = Host.CreateDefaultBuilder(args)
    .UseWindowsService(options =>
    {
        options.ServiceName = "ServizioAntiCopieMultiple";
    })
    .ConfigureServices(services =>
    {
        services.AddHostedService<PrintMonitorWorker>();
    })
    .Build();

host.Run();
