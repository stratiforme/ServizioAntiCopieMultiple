using System.Runtime.Versioning;
using System;
using System.ServiceProcess;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Threading;

[assembly: SupportedOSPlatform("windows")]

const string serviceName = "ServizioAntiCopieMultiple";

// Current UI language (two-letter code). Default to Italian
string currentLanguage = "it";

// Simple localization helper for this console tool (supports 'it' and 'en')
static class Localizer
{
    private static string _lang = "it";
    public static string Language { get => _lang; set { _lang = string.IsNullOrEmpty(value) ? "it" : value; } }
    private static readonly Dictionary<string, Dictionary<string, string>> _map = new()
    {
        ["it"] = new Dictionary<string,string>
        {
            ["ToolTitle"] = "Servizio Anti Copie Multiple - Tool di gestione",
            ["ServiceInstalled"] = "Servizio installato: {0}",
            ["ServiceStatus"] = "Stato servizio: {0}",
            ["CannotGetStatus"] = "Impossibile ottenere lo stato del servizio: {0}",
            ["ChooseAction"] = "Scegli un'azione:",
            ["InstallService"] = "1) Installa servizio",
            ["UninstallService"] = "1) Disinstalla servizio",
            ["StartService"] = "2) Avvia servizio come servizio di Windows",
            ["StopService"] = "3) Ferma servizio",
            ["RunConsole"] = "4) Esegui servizio in modalità console (--console) (utile per debug/diagnostica)",
            ["ConfigureManual"] = "5) Configura impostazioni stampa manualmente",
            ["Exit"] = "0) Esci",
            ["NotAdminRetry"] = "Impossibile ottenere i privilegi elevati automaticamente. Avvia il tool come amministratore.",
            ["InstallPromptExeNotFound"] = "Eseguibile del servizio non trovato automaticamente.",
            ["InstallAskExePath"] = "Inserisci il percorso completo dell'eseguibile del servizio (o premi invio per annullare): ",
            ["InstallExeNotSpecified"] = "Eseguibile non specificato o non trovato.",
            ["EventLogCreateWarning"] = "Warning: impossibile creare la sorgente EventLog: {0}",
            ["InstallSucceededAskConfigure"] = "Installazione riuscita. Vuoi configurare le impostazioni ora? (S/n): ",
            ["ConfigSavedAt"] = "Configurazione salvata in: {0}",
            ["InstallSuccess"] = "Installazione riuscita.",
            ["InstallFailed"] = "Installazione fallita: {0}",
            ["UninstallSuccess"] = "Disinstallazione riuscita.",
            ["UninstallFailed"] = "Disinstallazione fallita: {0}",
            ["StartSuccess"] = "Servizio avviato.",
            ["StartFailed"] = "Avvio fallito: {0}",
            ["StopSuccess"] = "Servizio fermato.",
            ["StopFailed"] = "Arresto fallito: {0}",
            ["RunConsoleStarted"] = "Servizio avviato in modalità console (PID {0}). Premi invio per terminare il processo e ritornare al menu.",
            ["PressEnterToContinue"] = "Premi invio per continuare...",
            ["PromptChoice"] = "Scelta: ",
            ["ConfigMenuTitle"] = "Configurazione PrintMonitor (premi invio per usare il valore predefinito)",
            ["EnableScannerPrompt"] = "Abilitare lo scanner in servizio?",
            ["SaveNetDumpPrompt"] = "Salvare dump diagnostici per stampanti di rete?",
            ["EnableNetCancelPrompt"] = "Consentire cancellazione automatica per stampanti di rete?",
            ["ScanIntervalPrompt"] = "Intervallo scanner (secondi)",
            ["JobAgePrompt"] = "Soglia età job (secondi)",
            ["SigWindowPrompt"] = "Finestra signature (secondi)",
            ["ConfigUpdated"] = "Configurazione aggiornata.",
            ["LanguagePrompt"] = "Scegli lingua / Choose language: 1) Italiano 2) English (default Italiano): ",
            ["LanguageSet"] = "Lingua impostata su: {0}"
        },
        ["en"] = new Dictionary<string,string>
        {
            ["ToolTitle"] = "Anti Multiple Copies Service - Management Tool",
            ["ServiceInstalled"] = "Service installed: {0}",
            ["ServiceStatus"] = "Service status: {0}",
            ["CannotGetStatus"] = "Unable to get service status: {0}",
            ["ChooseAction"] = "Choose an action:",
            ["InstallService"] = "1) Install service",
            ["UninstallService"] = "1) Uninstall service",
            ["StartService"] = "2) Start service (Windows service)",
            ["StopService"] = "3) Stop service",
            ["RunConsole"] = "4) Run service in console mode (--console) (useful for debug)",
            ["ConfigureManual"] = "5) Configure print settings manually",
            ["Exit"] = "0) Exit",
            ["NotAdminRetry"] = "Unable to automatically obtain elevated privileges. Run the tool as administrator.",
            ["InstallPromptExeNotFound"] = "Service executable not found automatically.",
            ["InstallAskExePath"] = "Enter full path to service executable (or press enter to cancel): ",
            ["InstallExeNotSpecified"] = "Executable not specified or not found.",
            ["EventLogCreateWarning"] = "Warning: unable to create EventLog source: {0}",
            ["InstallSucceededAskConfigure"] = "Installation succeeded. Configure settings now? (Y/n): ",
            ["ConfigSavedAt"] = "Configuration saved at: {0}",
            ["InstallSuccess"] = "Installation succeeded.",
            ["InstallFailed"] = "Installation failed: {0}",
            ["UninstallSuccess"] = "Uninstallation succeeded.",
            ["UninstallFailed"] = "Uninstallation failed: {0}",
            ["StartSuccess"] = "Service started.",
            ["StartFailed"] = "Start failed: {0}",
            ["StopSuccess"] = "Service stopped.",
            ["StopFailed"] = "Stop failed: {0}",
            ["RunConsoleStarted"] = "Service started in console mode (PID {0}). Press Enter to stop the process and return to the menu.",
            ["PressEnterToContinue"] = "Press Enter to continue...",
            ["PromptChoice"] = "Choice: ",
            ["ConfigMenuTitle"] = "PrintMonitor configuration (press enter to use default)",
            ["EnableScannerPrompt"] = "Enable scanner in service?",
            ["SaveNetDumpPrompt"] = "Save diagnostics dumps for network printers?",
            ["EnableNetCancelPrompt"] = "Allow automatic cancellation for network printers?",
            ["ScanIntervalPrompt"] = "Scanner interval (seconds)",
            ["JobAgePrompt"] = "Job age threshold (seconds)",
            ["SigWindowPrompt"] = "Signature window (seconds)",
            ["ConfigUpdated"] = "Configuration updated.",
            ["LanguagePrompt"] = "Choose language / Scegli lingua: 1) Italiano 2) English (default Italiano): ",
            ["LanguageSet"] = "Language set to: {0}"
        }
    };

    public static string T(string key)
    {
        var lang = _lang ?? "it";
        if (_map.TryGetValue(lang, out var dict) && dict.TryGetValue(key, out var val)) return val;
        // fallback to italian
        if (_map.TryGetValue("it", out var def) && def.TryGetValue(key, out var dval)) return dval;
        return key;
    }
}

string GetSharedConfigPath()
{
    var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple");
    Directory.CreateDirectory(dir);
    return Path.Combine(dir, "config.json");
}

void LoadLanguageFromConfig()
{
    try
    {
        var path = GetSharedConfigPath();
        if (!File.Exists(path)) return;
        var doc = System.Text.Json.JsonDocument.Parse(File.ReadAllText(path));
        if (doc.RootElement.TryGetProperty("Language", out var langEl))
        {
            var lang = langEl.GetString();
            if (!string.IsNullOrEmpty(lang))
            {
                currentLanguage = lang;
                Localizer.Language = lang;
            }
        }
    }
    catch { }
}

void PromptLanguage()
{
    Console.Write(Localizer.T("LanguagePrompt"));
    var sel = Console.ReadLine();
    if (string.IsNullOrWhiteSpace(sel) || sel.Trim() == "1")
    {
        currentLanguage = "it";
    }
    else
    {
        currentLanguage = "en";
    }
    Localizer.Language = currentLanguage;
    Console.WriteLine(Localizer.T("LanguageSet"), currentLanguage);

    // Persist language choice to shared config so the service and future runs pick it up
    try
    {
        var path = GetSharedConfigPath();
        Dictionary<string, object?> root = new();
        if (File.Exists(path))
        {
            try
            {
                root = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object?>>(File.ReadAllText(path)) ?? new();
            }
            catch { root = new(); }
        }

        root["Language"] = currentLanguage;

        var opts = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(root, opts));
    }
    catch { }
}

// Try load language before showing UI
LoadLanguageFromConfig();
if (string.IsNullOrEmpty(Localizer.Language)) PromptLanguage();

bool IsAdministrator()
{
    try
    {
        using var id = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(id);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }
    catch
    {
        return false;
    }
}

bool TryRestartAsAdmin()
{
    try
    {
        var exe = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        if (string.IsNullOrEmpty(exe)) return false;
        var args = string.Join(' ', Environment.GetCommandLineArgs().Skip(1).Select(a => a.Contains(' ') ? '"' + a + '"' : a));
        var psi = new ProcessStartInfo(exe, args)
        {
            UseShellExecute = true,
            Verb = "runas"
        };
        Process.Start(psi);
        return true;
    }
    catch
    {
        return false;
    }
}

void EnsureElevated()
{
    if (!IsAdministrator())
    {
        // Attempt to relaunch elevated automatically. Note: this will trigger the UAC prompt; cannot be bypassed programmatically.
        if (TryRestartAsAdmin())
        {
            // Exit current non-elevated instance to allow elevated one to run
            Environment.Exit(0);
        }
        else
        {
            Console.WriteLine(Localizer.T("NotAdminRetry"));
            Thread.Sleep(1500);
        }
    }
}

void ShowMenu()
{
    Console.Clear();
    Console.WriteLine(Localizer.T("ToolTitle") + "\n");

    bool isInstalled = ServiceController.GetServices().Any(s => s.ServiceName.Equals(serviceName, StringComparison.OrdinalIgnoreCase));
    Console.WriteLine(string.Format(Localizer.T("ServiceInstalled"), isInstalled));
    if (isInstalled)
    {
        try
        {
            using var sc = new ServiceController(serviceName);
            Console.WriteLine(string.Format(Localizer.T("ServiceStatus"), sc.Status));
        }
        catch (Exception ex)
        {
            Console.WriteLine(string.Format(Localizer.T("CannotGetStatus"), ex.Message));
        }
    }

    Console.WriteLine('\n' + Localizer.T("ChooseAction"));
    // Make choices stable and logical
    if (!isInstalled)
    {
        Console.WriteLine(Localizer.T("InstallService"));
         // no unrelated utilities here
        Console.WriteLine(Localizer.T("Exit"));
    }
    else
    {
        Console.WriteLine(Localizer.T("UninstallService"));
        Console.WriteLine(Localizer.T("StartService"));
        Console.WriteLine(Localizer.T("StopService"));
        Console.WriteLine(Localizer.T("RunConsole"));
        Console.WriteLine(Localizer.T("ConfigureManual"));
        Console.WriteLine(Localizer.T("Exit"));
    }
}

string? FindServiceExe()
{
    // First, look in the same folder as the tool
    string baseDir = AppContext.BaseDirectory;
    string candidate = Path.Combine(baseDir, "ServizioAntiCopieMultiple.exe");
    if (File.Exists(candidate)) return candidate;

    // Also consider same folder with different casing
    var exeInBase = Directory.EnumerateFiles(baseDir, "ServizioAntiCopieMultiple*.exe").FirstOrDefault();
    if (exeInBase != null) return exeInBase;

    // Check common publish/build locations relative to repo layout
    var fallbacks = new[]
    {
        Path.GetFullPath(Path.Combine(baseDir, "..", "ServizioAntiCopieMultiple", "bin", "Release", "net10.0", "publish", "ServizioAntiCopieMultiple.exe")),
        Path.GetFullPath(Path.Combine(baseDir, "..", "ServizioAntiCopieMultiple", "bin", "Release", "net10.0", "ServizioAntiCopieMultiple.exe")),
        Path.GetFullPath(Path.Combine(baseDir, "..", "ServizioAntiCopieMultiple", "bin", "Debug", "net10.0", "publish", "ServizioAntiCopieMultiple.exe")),
        Path.GetFullPath(Path.Combine(baseDir, "..", "ServizioAntiCopieMultiple", "bin", "Debug", "net10.0", "ServizioAntiCopieMultiple.exe")),
        // Also check artifacts publish folder used by CI
        Path.GetFullPath(Path.Combine(baseDir, "..", "artifacts", "publish", "ServizioAntiCopieMultiple.exe")),
        Path.GetFullPath(Path.Combine(baseDir, "..", "..", "artifacts", "publish", "ServizioAntiCopieMultiple.exe"))
    };

    foreach (var f in fallbacks)
    {
        if (File.Exists(f)) return f;
    }

    return null;
}

bool TryInstallService(out string message)
{
    if (!IsAdministrator())
    {
        message = Localizer.T("NotAdminRetry");
        return false;
    }

    try
    {
        string? exePath = FindServiceExe();
        if (string.IsNullOrEmpty(exePath))
        {
            Console.WriteLine(Localizer.T("InstallPromptExeNotFound"));
            Console.Write(Localizer.T("InstallAskExePath"));
            var input = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(input) || !File.Exists(input))
            {
                message = Localizer.T("InstallExeNotSpecified");
                return false;
            }
            exePath = input.Trim('"');
        }

        // sc create requires cmd invocation
        var psi = new ProcessStartInfo("sc.exe", $"create {serviceName} binPath= \"{exePath}\" start= auto DisplayName= \"Servizio Anti Copie Multiple\"")
        {
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var p = Process.Start(psi)!;
        string output = p.StandardOutput.ReadToEnd();
        p.WaitForExit();
        message = output;

        bool created = p.ExitCode == 0 || output.Contains("CreateService SUCCESS");
        if (created)
        {
            // Try to create EventLog source so the service can write to Application log
            try
            {
                if (!EventLog.SourceExists(serviceName))
                {
                    EventLog.CreateEventSource(new EventSourceCreationData(serviceName, "Application"));
                    message += "\nEventLog source created.";
                }
            }
            catch (Exception ex)
            {
                message += "\n" + string.Format(Localizer.T("EventLogCreateWarning"), ex.Message);
            }

            // After successful install, offer to configure settings now and save to ProgramData
            Console.Write(Localizer.T("InstallSucceededAskConfigure"));
            var ans = Console.ReadLine();
            if (string.IsNullOrWhiteSpace(ans) || ans.Trim().Equals("s", StringComparison.OrdinalIgnoreCase) || ans.Trim().Equals("y", StringComparison.OrdinalIgnoreCase))
             {
                 var cfg = PromptPrintMonitorSettings();
                 SaveConfigToCommonAppData(cfg);
             }
         }

         return created;
     }
     catch (Exception ex)
     {
         message = ex.Message;
         return false;
     }
}

bool TryUninstallService(out string message)
{
    if (!IsAdministrator())
    {
        message = Localizer.T("NotAdminRetry");
        return false;
    }

    try
    {
        var psi = new ProcessStartInfo("sc.exe", $"delete {serviceName}")
        {
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var p = Process.Start(psi)!;
        string output = p.StandardOutput.ReadToEnd();
        p.WaitForExit();
        message = output;

        bool deleted = p.ExitCode == 0 || output.Contains("DeleteService SUCCESS");
        if (deleted)
        {
            // Try to remove EventLog source
            try
            {
                if (EventLog.SourceExists(serviceName))
                {
                    EventLog.DeleteEventSource(serviceName);
                    message += "\nEventLog source removed.";
                }
            }
            catch (Exception ex)
            {
                message += "\nWarning: impossibile rimuovere la sorgente EventLog: " + ex.Message;
            }
        }

        return deleted;
    }
    catch (Exception ex)
    {
        message = ex.Message;
        return false;
    }
}

bool TryStartService(out string message)
{
    if (!IsAdministrator())
    {
        message = Localizer.T("NotAdminRetry");
        return false;
    }

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
    if (!IsAdministrator())
    {
        message = Localizer.T("NotAdminRetry");
        return false;
    }

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

bool TryRunServiceConsole(out string message)
{
    // Start the service executable with --console to run in foreground. This does not require admin
    try
    {
        var exe = FindServiceExe();
        if (string.IsNullOrEmpty(exe))
        {
            message = "Eseguibile del servizio non trovato.";
            return false;
        }

        var psi = new ProcessStartInfo(exe, "--console")
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = false
        };

        var proc = Process.Start(psi);
        if (proc == null)
        {
            message = "Impossibile avviare il processo.";
            return false;
        }

        // Stream output to current console
        proc.BeginOutputReadLine();
        proc.BeginErrorReadLine();
        proc.OutputDataReceived += (s, e) => { if (e.Data != null) Console.WriteLine(e.Data); };
        proc.ErrorDataReceived += (s, e) => { if (e.Data != null) Console.Error.WriteLine(e.Data); };

        message = $"Servizio avviato in modalità console (PID {proc.Id}). Premi invio per terminare il processo e ritornare al menu.";
        return true;
    }
    catch (Exception ex)
    {
        message = ex.Message;
        return false;
    }
}

void SaveConfigToCommonAppData(Dictionary<string, object> printMonitorSettings)
{
    try
    {
        var root = new Dictionary<string, object?>();

        // preserve existing Language if present
        string path = GetSharedConfigPath();
        if (File.Exists(path))
        {
            try
            {
                var existing = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object?>>(File.ReadAllText(path));
                if (existing != null && existing.TryGetValue("Language", out var langVal))
                {
                    root["Language"] = langVal;
                }
            }
            catch { }
        }

        root["PrintMonitor"] = printMonitorSettings;

        var opts = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(root, opts));
        Console.WriteLine(string.Format(Localizer.T("ConfigSavedAt"), path));
    }
    catch (Exception ex)
    {
        Console.WriteLine("Errore salvando la configurazione: " + ex.Message);
    }
}

Dictionary<string, object> PromptPrintMonitorSettings()
{
    var settings = new Dictionary<string, object>();

    bool ReadBool(string prompt, bool defaultVal)
    {
        // Show a clear prompt with the default highlighted: use "S/n" when default is Yes, "s/N" when default is No
        string hint = defaultVal ? "S/n" : "s/N";
        Console.Write(prompt + $" ({hint}): ");
        var a = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(a)) return defaultVal;
        var t = a.Trim();
        // Accept common variants: 's', 'si', 'y' for yes; 'n', 'no' for no
        if (t.Equals("s", StringComparison.OrdinalIgnoreCase) || t.Equals("si", StringComparison.OrdinalIgnoreCase) || t.Equals("y", StringComparison.OrdinalIgnoreCase)) return true;
        if (t.Equals("n", StringComparison.OrdinalIgnoreCase) || t.Equals("no", StringComparison.OrdinalIgnoreCase)) return false;
        // Fallback to default if input unrecognized
        return defaultVal;
    }

    int ReadInt(string prompt, int defaultVal)
    {
        Console.Write(prompt + $" (default {defaultVal}): ");
        var t = Console.ReadLine();
        if (int.TryParse(t, out var v)) return v;
        return defaultVal;
    }

    Console.WriteLine(Localizer.T("ConfigMenuTitle"));
    bool enableScanner = ReadBool(Localizer.T("EnableScannerPrompt"), false);
    bool saveNetworkDumps = ReadBool(Localizer.T("SaveNetDumpPrompt"), true);
    bool enableNetworkCancel = ReadBool(Localizer.T("EnableNetCancelPrompt"), false);
    int scanInterval = ReadInt(Localizer.T("ScanIntervalPrompt"), 5);
    int jobAge = ReadInt(Localizer.T("JobAgePrompt"), 30);
    int sigWindow = ReadInt(Localizer.T("SigWindowPrompt"), 10);

    settings["EnableScannerInService"] = enableScanner;
    settings["SaveNetworkDumps"] = saveNetworkDumps;
    settings["EnableNetworkCancellation"] = enableNetworkCancel;
    settings["ScanIntervalSeconds"] = scanInterval;
    settings["JobAgeThresholdSeconds"] = jobAge;
    settings["SignatureWindowSeconds"] = sigWindow;

    return settings;
}

// Ensure elevated privileges at startup when possible
EnsureElevated();

while (true)
{
    ShowMenu();
    Console.Write(Localizer.T("PromptChoice"));
    var ch = Console.ReadLine();
    if (string.IsNullOrWhiteSpace(ch)) continue;
    if (ch == "0") break;

    string resultMsg = string.Empty;
    bool ok = false;
    if (ch == "1")
    {
        bool isInstalled = ServiceController.GetServices().Any(s => s.ServiceName.Equals(serviceName, StringComparison.OrdinalIgnoreCase));
        if (!isInstalled)
        {
            ok = TryInstallService(out resultMsg);
            Console.WriteLine(ok ? Localizer.T("InstallSuccess") : Localizer.T("InstallFailed"), resultMsg);
        }
        else
        {
            ok = TryUninstallService(out resultMsg);
            Console.WriteLine(ok ? Localizer.T("UninstallSuccess") : Localizer.T("UninstallFailed"), resultMsg);
        }
    }
    else if (ch == "5")
    {
        // Manual configuration helper
        var cfg = PromptPrintMonitorSettings();
        SaveConfigToCommonAppData(cfg);
        Console.WriteLine(Localizer.T("ConfigUpdated"));
        ok = true;
    }
    else if (ch == "2")
    {
        ok = TryStartService(out resultMsg);
        Console.WriteLine(ok ? resultMsg : Localizer.T("StartFailed"), resultMsg);
    }
    else if (ch == "3")
    {
        ok = TryStopService(out resultMsg);
        Console.WriteLine(ok ? resultMsg : Localizer.T("StopFailed"), resultMsg);
    }
    else if (ch == "4")
    {
        ok = TryRunServiceConsole(out resultMsg);
        Console.WriteLine(ok ? resultMsg : "Avvio in console fallito: " + resultMsg);
        if (ok)
        {
            Console.WriteLine("Premi invio per terminare il processo in foreground e tornare al menu...");
            Console.ReadLine();
            // Attempt to kill any child process started with --console
            // Note: for simplicity we attempt to find processes with the same exe name and kill the most recent one
            try
            {
                var exe = Path.GetFileName(FindServiceExe() ?? string.Empty);
                if (!string.IsNullOrEmpty(exe))
                {
                    var procs = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(exe));
                    foreach (var p in procs.OrderByDescending(p => p.StartTime).Take(1))
                    {
                        try { p.Kill(); } catch { }
                    }
                }
            }
            catch { }
        }
    }

    Console.WriteLine(Localizer.T("PressEnterToContinue"));
    Console.ReadLine();
}
