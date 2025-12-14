using System;
using System.ServiceProcess;
using System.Diagnostics;
using System.IO;
using System.Linq;

const string serviceName = "ServizioAntiCopieMultiple";

void ShowMenu()
{
    Console.Clear();
    Console.WriteLine("Servizio Anti Copie Multiple - Tool di gestione\n");

    bool isInstalled = ServiceController.GetServices().Any(s => s.ServiceName.Equals(serviceName, StringComparison.OrdinalIgnoreCase));
    Console.WriteLine($"Servizio installato: {isInstalled}");
    if (isInstalled)
    {
        try
        {
            using var sc = new ServiceController(serviceName);
            Console.WriteLine($"Stato servizio: {sc.Status}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Impossibile ottenere lo stato del servizio: {ex.Message}");
        }
    }

    Console.WriteLine('\n' + "Scegli un'azione:");
    if (!isInstalled)
    {
        Console.WriteLine("1) Installa servizio");
    }
    else
    {
        Console.WriteLine("1) Disinstalla servizio");
        Console.WriteLine("2) Avvia servizio");
        Console.WriteLine("3) Ferma servizio");
    }
    Console.WriteLine("0) Esci");
}

bool TryInstallService(out string message)
{
    try
    {
        string exePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "ServizioAntiCopieMultiple", "bin", "Release", "net10.0", "publish", "ServizioAntiCopieMultiple.exe"));
        // if the publish folder does not exist, fallback to build output
        if (!File.Exists(exePath))
        {
            exePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "ServizioAntiCopieMultiple", "bin", "Release", "net10.0", "ServizioAntiCopieMultiple.exe"));
        }

        if (!File.Exists(exePath))
        {
            message = $"Eseguibile non trovato: {exePath} - compila e pubblica prima di installare.";
            return false;
        }

        // sc create requires cmd invocation
        var psi = new ProcessStartInfo("sc.exe", $"create {serviceName} binPath= \"{exePath}\" start= auto DisplayName= \"Servizio Anti Copie Multiple\"")
        {
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var p = Process.Start(psi);
        string output = p!.StandardOutput.ReadToEnd();
        p.WaitForExit();
        message = output;
        return p.ExitCode == 0 || output.Contains("CreateService SUCCESS");
    }
    catch (Exception ex)
    {
        message = ex.Message;
        return false;
    }
}

bool TryUninstallService(out string message)
{
    try
    {
        var psi = new ProcessStartInfo("sc.exe", $"delete {serviceName}")
        {
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var p = Process.Start(psi);
        string output = p!.StandardOutput.ReadToEnd();
        p.WaitForExit();
        message = output;
        return p.ExitCode == 0 || output.Contains("DeleteService SUCCESS");
    }
    catch (Exception ex)
    {
        message = ex.Message;
        return false;
    }
}

bool TryStartService(out string message)
{
    try
    {
        using var sc = new ServiceController(serviceName);
        if (sc.Status == ServiceControllerStatus.Running)
        {
            message = "Il servizio è già in esecuzione.";
            return true;
        }
        sc.Start();
        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(15));
        message = "Servizio avviato.";
        return true;
    }
    catch (Exception ex)
    {
        message = ex.Message;
        return false;
    }
}

bool TryStopService(out string message)
{
    try
    {
        using var sc = new ServiceController(serviceName);
        if (sc.Status == ServiceControllerStatus.Stopped)
        {
            message = "Il servizio è già fermo.";
            return true;
        }
        sc.Stop();
        sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(15));
        message = "Servizio fermato.";
        return true;
    }
    catch (Exception ex)
    {
        message = ex.Message;
        return false;
    }
}

while (true)
{
    ShowMenu();
    Console.Write("Scelta: ");
    var ch = Console.ReadLine();
    if (string.IsNullOrWhiteSpace(ch)) continue;
    if (ch == "0") break;

    string resultMsg;
    bool ok;
    if (ch == "1")
    {
        bool isInstalled = ServiceController.GetServices().Any(s => s.ServiceName.Equals(serviceName, StringComparison.OrdinalIgnoreCase));
        if (!isInstalled)
        {
            ok = TryInstallService(out resultMsg);
            Console.WriteLine(ok ? "Installazione riuscita." : "Installazione fallita: " + resultMsg);
        }
        else
        {
            ok = TryUninstallService(out resultMsg);
            Console.WriteLine(ok ? "Disinstallazione riuscita." : "Disinstallazione fallita: " + resultMsg);
        }
    }
    else if (ch == "2")
    {
        ok = TryStartService(out resultMsg);
        Console.WriteLine(ok ? resultMsg : "Avvio fallito: " + resultMsg);
    }
    else if (ch == "3")
    {
        ok = TryStopService(out resultMsg);
        Console.WriteLine(ok ? resultMsg : "Arresto fallito: " + resultMsg);
    }

    Console.WriteLine("Premi invio per continuare...");
    Console.ReadLine();
}
