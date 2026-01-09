using System.Diagnostics;
using System.IO;
using System.Runtime.Versioning;
using System.Security.Principal;
using System.ServiceProcess;

[assembly: SupportedOSPlatform("windows")]

namespace GestioneSacm
{
    internal static class Program
    {
        private const string ServiceName = "ServizioAntiCopieMultiple";

        private static void Main(string[] args)
        {
            EnsureElevated();

            while (true)
            {
                ShowMenu();
                Console.Write("Scelta: ");
                var ch = Console.ReadLine();
                if (string.IsNullOrWhiteSpace(ch)) continue;
                if (ch == "0") break;

                string resultMsg = string.Empty;
                bool ok = false;
                if (ch == "1")
                {
                    bool isInstalled = ServiceController.GetServices().Any(s => s.ServiceName.Equals(ServiceName, StringComparison.OrdinalIgnoreCase));
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
                else if (ch == "4")
                {
                    ok = TryRunServiceConsole(out resultMsg);
                    Console.WriteLine(ok ? resultMsg : "Avvio in console fallito: " + resultMsg);
                    if (ok)
                    {
                        Console.WriteLine("Premi invio per terminare il processo in foreground e tornare al menu...");
                        Console.ReadLine();
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

                Console.WriteLine("Premi invio per continuare...");
                Console.ReadLine();
            }
        }

        private static bool IsAdministrator()
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

        private static bool TryRestartAsAdmin()
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

        private static void EnsureElevated()
        {
            if (!IsAdministrator())
            {
                // Tentativo automatico di elevazione senza prompt "S/n"
                if (TryRestartAsAdmin())
                {
                    Environment.Exit(0);
                }
                else
                {
                    Console.WriteLine("Impossibile ottenere privilegi elevati automaticamente. Il tool continuerà in modalità non elevata.");
                    Thread.Sleep(1000);
                }
            }
        }

        private static void ShowMenu()
        {
            Console.Clear();
            Console.WriteLine("Servizio Anti Copie Multiple - Tool di gestione\n");

            bool isInstalled = ServiceController.GetServices().Any(s => s.ServiceName.Equals(ServiceName, StringComparison.OrdinalIgnoreCase));
            Console.WriteLine($"Servizio installato: {isInstalled}");

            if (isInstalled)
            {
                try
                {
                    using var sc = new ServiceController(ServiceName);
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
                Console.WriteLine("0) Esci");
            }
            else
            {
                Console.WriteLine("1) Disinstalla servizio");
                Console.WriteLine("2) Avvia servizio come servizio di Windows");
                Console.WriteLine("3) Ferma servizio");
                Console.WriteLine("4) Esegui servizio in modalità console (--console) (utile per debug/diagnostica)");
                Console.WriteLine("0) Esci");
            }
        }

        private static string? FindServiceExe()
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

        private static bool TryInstallService(out string message)
        {
            if (!IsAdministrator())
            {
                message = "Operazione richiede privilegi amministrativi. Riavvia il tool come amministratore.";
                return false;
            }

            try
            {
                string? exePath = FindServiceExe();
                if (string.IsNullOrEmpty(exePath))
                {
                    Console.WriteLine("Eseguibile del servizio non trovato automaticamente.");
                    Console.Write("Inserisci il percorso completo dell'eseguibile del servizio (o premi invio per annullare): ");
                    var input = Console.ReadLine();
                    if (string.IsNullOrWhiteSpace(input) || !File.Exists(input))
                    {
                        message = "Eseguibile non specificato o non trovato.";
                        return false;
                    }
                    exePath = input.Trim('"');
                }

                // sc create requires cmd invocation
                var psi = new ProcessStartInfo("sc.exe", $"create {ServiceName} binPath= \"{exePath}\" start= auto DisplayName= \"Servizio Anti Copie Multiple\"")
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
                        if (!EventLog.SourceExists(ServiceName))
                        {
                            EventLog.CreateEventSource(new EventSourceCreationData(ServiceName, "Application"));
                            message += "\nEventLog source created.";
                        }
                    }
                    catch (Exception ex)
                    {
                        message += "\nWarning: impossibile creare la sorgente EventLog: " + ex.Message;
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

        private static bool TryUninstallService(out string message)
        {
            if (!IsAdministrator())
            {
                message = "Operazione richiede privilegi amministrativi. Riavvia il tool come amministratore.";
                return false;
            }

            try
            {
                var psi = new ProcessStartInfo("sc.exe", $"delete {ServiceName}")
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
                        if (EventLog.SourceExists(ServiceName))
                        {
                            EventLog.DeleteEventSource(ServiceName);
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

        private static bool TryStartService(out string message)
        {
            if (!IsAdministrator())
            {
                message = "Operazione richiede privilegi amministrativi. Riavvia il tool come amministratore.";
                return false;
            }

            try
            {
                using var sc = new ServiceController(ServiceName);
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

        private static bool TryStopService(out string message)
        {
            if (!IsAdministrator())
            {
                message = "Operazione richiede privilegi amministrativi. Riavvia il tool come amministratore.";
                return false;
            }

            try
            {
                using var sc = new ServiceController(ServiceName);
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

        private static bool TryRunServiceConsole(out string message)
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
    }
}