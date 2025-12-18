using System;
using System.Collections.Concurrent;
using System.IO;
using System.Management;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System.Runtime.Versioning;
using System.Collections.Generic;
using System.Text.Json;

namespace ServizioAntiCopieMultiple
{
    [SupportedOSPlatform("windows")]
    public class PrintMonitorWorker : BackgroundService
    {
        private readonly ILogger<PrintMonitorWorker> _logger;
        private ManagementEventWatcher? _printJobWatcher;
        private FileSystemWatcher? _responseWatcher;
        private FileSystemWatcher? _simulatorWatcher;
        private readonly ConcurrentDictionary<string, bool> _observedJobs = new();
        private readonly string _responsesDir;
        private readonly string _simulatorDir;
        private readonly PrintJobProcessor _processor = new();

        public PrintMonitorWorker(ILogger<PrintMonitorWorker> logger)
        {
            _logger = logger;
            _responsesDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple", "responses");
            _simulatorDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple", "simulator");
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
                _logger.LogInformation("PrintMonitorWorker starting. ProcessId={Pid}, User={User}", Environment.ProcessId, Environment.UserName);

                // Ensure responses directory exists for user click simulation
                Directory.CreateDirectory(_responsesDir);
                _logger.LogInformation("Responses directory ensured at {dir}", _responsesDir);

                // Ensure simulator directory exists for local testing (drop JSON files to simulate print jobs)
                Directory.CreateDirectory(_simulatorDir);
                _logger.LogInformation("Simulator directory ensured at {dir}", _simulatorDir);

                // FileSystemWatcher to detect simulator JSON files
                _simulatorWatcher = new FileSystemWatcher(_simulatorDir, "*.json")
                {
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
                };
                _simulatorWatcher.Created += OnSimulatorFileCreated;
                _simulatorWatcher.EnableRaisingEvents = true;
                _logger.LogInformation("Simulator FileSystemWatcher started watching {dir}", _simulatorDir);

                // FileSystemWatcher to detect user OK responses (external UI/tool can drop files named <jobId>.ok)
                _responseWatcher = new FileSystemWatcher(_responsesDir, "*.ok")
                {
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
                };
                _responseWatcher.Created += OnResponseFileCreated;
                _responseWatcher.EnableRaisingEvents = true;
                _logger.LogInformation("Response FileSystemWatcher started watching {dir}", _responsesDir);

                // Start WMI connect in background to avoid blocking
                _ = Task.Run(async () => await EnsureWmiWatcherAsync(stoppingToken).ConfigureAwait(false));

                _logger.LogInformation("PrintMonitorWorker started and monitoring print jobs. Responses dir: {dir}", _responsesDir);

                while (!stoppingToken.IsCancellationRequested)
                {
                    await Task.Delay(1000, stoppingToken).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException)
            {
                // shutdown
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unhandled exception in ExecuteAsync");
                throw;
            }
            finally
            {
                await _processor.StopAsync().ConfigureAwait(false);
                DisposeWatchers();
            }
        }

        // Public helper to simulate a user clicking OK (creates a .ok file)
        public void SimulateUserOk(string jobId)
        {
            try
            {
                string responseFile = Path.Combine(_responsesDir, jobId + ".ok");
                File.WriteAllText(responseFile, string.Empty);
                _logger.LogInformation("SimulateUserOk: created {File}", responseFile);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "SimulateUserOk: failed to create OK file for {JobId}", jobId);
            }
        }

        // Simulator file format (JSON) example:
        // { "Name": "Printer, 123", "Document": "Doc.pdf", "Owner": "user", "Copies": 5, "Path": "\\\\\\?" }
        private async void OnSimulatorFileCreated(object? sender, FileSystemEventArgs e)
        {
            // small delay to allow file write to complete
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    if (File.Exists(e.FullPath)) break;
                }
                catch { }
                await Task.Delay(100).ConfigureAwait(false);
            }

            try
            {
                string content = string.Empty;
                try
                {
                    content = File.ReadAllText(e.FullPath);
                }
                catch
                {
                    // try again once
                    await Task.Delay(100).ConfigureAwait(false);
                    content = File.ReadAllText(e.FullPath);
                }

                var info = JsonSerializer.Deserialize<Dictionary<string, object?>>(content);
                if (info == null)
                {
                    _logger.LogWarning("Simulator: could not parse JSON from {File}", e.FullPath);
                    return;
                }

                var jobInfo = new PrintJobInfo
                {
                    Name = info.TryGetValue("Name", out var n) ? n?.ToString() : null,
                    Document = info.TryGetValue("Document", out var d) ? d?.ToString() ?? string.Empty : string.Empty,
                    Owner = info.TryGetValue("Owner", out var o) ? o?.ToString() ?? string.Empty : string.Empty,
                    Copies = info.TryGetValue("Copies", out var c) && int.TryParse(c?.ToString(), out var ci) ? ci : 1,
                    Path = info.TryGetValue("Path", out var p) ? p?.ToString() ?? string.Empty : string.Empty
                };

                _logger.LogInformation("Simulator: invoking simulated print job: Name={Name}, Document={Document}, Owner={Owner}, Copies={Copies}", jobInfo.Name, jobInfo.Document, jobInfo.Owner, jobInfo.Copies);
                ProcessPrintJobInfo(jobInfo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Simulator: error processing simulator file {File}", e.FullPath);
            }
            finally
            {
                // delete simulator file
                try { File.Delete(e.FullPath); } catch { }
            }
        }

        private record PrintJobInfo
        {
            public string? Name { get; init; }
            public string Document { get; init; } = string.Empty;
            public string Owner { get; init; } = string.Empty;
            public int Copies { get; init; } = 1;
            public string Path { get; init; } = string.Empty;
        }

        private async Task EnsureWmiWatcherAsync(CancellationToken token)
        {
            try
            {
                // Use connection options with privileges enabled as WMI print job operations sometimes require them
                var connOptions = new ConnectionOptions
                {
                    Impersonation = ImpersonationLevel.Impersonate,
                    EnablePrivileges = true,
                    Timeout = ManagementOptions.InfiniteTimeout
                };

                string query = "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_PrintJob'";
                var wql = new WqlEventQuery(query);

                // Try creating and connecting a scoped watcher inside a guarded block.
                try
                {
                    // Use verbatim literal for clarity
                    var scope = new ManagementScope(@"\\.\root\cimv2", connOptions);

                    scope.Connect();
                    _logger.LogInformation("WMI scope connected to {Namespace}", scope.Path?.Path);

                    // Create watcher bound to the connected scope
                    _printJobWatcher = new ManagementEventWatcher(scope, wql);
                    _printJobWatcher.EventArrived += OnPrintJobArrived;
                    _printJobWatcher.Stopped += OnWatcherStopped;

                    try
                    {
                        _printJobWatcher.Start();
                        _logger.LogInformation("ManagementEventWatcher started with query: {Query}", query);
                    }
                    catch (Exception startEx)
                    {
                        _logger.LogError(startEx, "Failed to start ManagementEventWatcher bound to scope");
                        // try fallback to an unbound watcher
                        try
                        {
                            _printJobWatcher?.Dispose();
                            _printJobWatcher = new ManagementEventWatcher(wql);
                            _printJobWatcher.EventArrived += OnPrintJobArrived;
                            _printJobWatcher.Stopped += OnWatcherStopped;
                            _printJobWatcher.Start();
                            _logger.LogInformation("ManagementEventWatcher started (fallback) with query: {Query}", query);
                        }
                        catch (Exception fbEx)
                        {
                            _logger.LogError(fbEx, "Failed to start fallback ManagementEventWatcher");
                        }
                    }
                }
                catch (Exception scopedEx)
                {
                    // Construction of ManagementScope or Connect() failed -> try unscoped fallback.
                    _logger.LogWarning(scopedEx, "Failed creating/connecting WMI scope; attempting unscoped ManagementEventWatcher fallback");

                    try
                    {
                        _printJobWatcher = new ManagementEventWatcher(wql);
                        _printJobWatcher.EventArrived += OnPrintJobArrived;
                        _printJobWatcher.Stopped += OnWatcherStopped;
                        _printJobWatcher.Start();
                        _logger.LogInformation("ManagementEventWatcher started (unscoped fallback) with query: {Query}", query);
                    }
                    catch (Exception fbEx)
                    {
                        _logger.LogError(fbEx, "Fallback unscoped ManagementEventWatcher also failed");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error during WMI watcher setup");
            }
        }

        private void OnResponseFileCreated(object? sender, FileSystemEventArgs e)
        {
            try
            {
                string jobId = Path.GetFileNameWithoutExtension(e.Name)!;
                _logger.LogInformation("UserClickedOk: received OK response for job {JobId} (file: {FileName})", jobId, e.FullPath);

                _observedJobs.TryAdd(jobId, true);

                // Schedule async delete via processor to avoid blocking FS event thread
                _ = _processor.EnqueueAsync(async () =>
                {
                    await TryDeleteFileWithRetriesAsync(e.FullPath).ConfigureAwait(false);
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error handling response file {FileName}", e.FullPath);
            }
        }

        private static async Task TryDeleteFileWithRetriesAsync(string path, int retries = 3)
        {
            if (string.IsNullOrEmpty(path)) return;
            for (int i = 0; i < retries; i++)
            {
                try
                {
                    if (File.Exists(path))
                        File.Delete(path);
                    return;
                }
                catch
                {
                    await Task.Delay(100).ConfigureAwait(false);
                }
            }
        }

        private void OnPrintJobArrived(object? sender, EventArrivedEventArgs e)
        {
            try
            {
                try
                {
                    var ev = e.NewEvent;
                    if (_logger.IsEnabled(Microsoft.Extensions.Logging.LogLevel.Debug))
                        _logger.LogDebug("WMI EventArrived invoked. EventClass: {EventClass}, TimeGenerated: {Time}", ev?.ClassPath?.ClassName, ev?["TIME_CREATED"]);
                }
                catch (Exception diagEx)
                {
                    _logger.LogDebug(diagEx, "Failed to read basic event metadata");
                }

                var target = (ManagementBaseObject?)e.NewEvent?["TargetInstance"];
                if (target == null)
                {
                    _logger.LogWarning("Print job event arrived but TargetInstance was null");
                    return;
                }

                // Dump all properties for debugging
                try
                {
                    DumpTargetProperties(target);
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Failed dumping TargetInstance properties");
                }

                string? name = target["Name"]?.ToString();
                string document = target["Document"]?.ToString() ?? string.Empty;
                string owner = target["Owner"]?.ToString() ?? string.Empty;

                try
                {
                    if (_logger.IsEnabled(Microsoft.Extensions.Logging.LogLevel.Debug))
                    {
                        string path = target["__PATH"]?.ToString() ?? string.Empty;
                        _logger.LogDebug("Print job properties: Name={Name}, Document={Document}, Owner={Owner}, __PATH={Path}", name, document, owner, path);
                    }
                }
                catch (Exception dbgEx)
                {
                    _logger.LogDebug(dbgEx, "Error reading target debug properties");
                }

                string jobId = PrintJobParser.ParseJobId(name);
                int copies = PrintJobParser.GetCopiesFromManagementObject(target);

                // Build PrintJobInfo and hand off to common processor
                var info = new PrintJobInfo
                {
                    Name = name,
                    Document = document,
                    Owner = owner,
                    Copies = copies,
                    Path = target["__PATH"]?.ToString() ?? string.Empty
                };

                if (info.Copies > 1)
                {
                    ProcessPrintJobInfo(info);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing print job event");
            }
        }

        private void DumpTargetProperties(ManagementBaseObject target)
        {
            try
            {
                var props = new Dictionary<string, string?>();
                foreach (PropertyData p in target.Properties)
                {
                    try
                    {
                        var val = p.Value?.ToString();
                        props[p.Name] = val;
                    }
                    catch (Exception ex)
                    {
                        props[p.Name] = $"<error: {ex.Message}>";
                    }
                }

                if (_logger.IsEnabled(Microsoft.Extensions.Logging.LogLevel.Debug))
                {
                    _logger.LogDebug("TargetInstance properties dump: {Props}", JsonSerializer.Serialize(props));
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Failed to dump TargetInstance properties");
            }
        }

        private void ProcessPrintJobInfo(PrintJobInfo info)
        {
            string jobId = PrintJobParser.ParseJobId(info.Name);
            _logger.LogInformation("DetectedMultiCopyPrintJob: JobId={JobId}, Document={Document}, Owner={Owner}, Copies={Copies}", jobId, info.Document, info.Owner, info.Copies);
            _logger.LogInformation("NotificationSent: Notification sent to user for job {JobId}", jobId);

            // Log expected OK response filename and path so client/UI can match exactly
            try
            {
                string expectedFileName = jobId + ".ok";
                _logger.LogInformation("ExpectedOk: create file {FileName} in {Dir} to allow job {JobId} to proceed", expectedFileName, _responsesDir, jobId);
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Failed to log expected OK filename");
            }

            if (_observedJobs.ContainsKey(jobId))
            {
                _logger.LogInformation("UserClickedOk: OK response already observed for job {JobId}; allowing job to proceed.", jobId);
                return;
            }

            string responseFile = Path.Combine(_responsesDir, jobId + ".ok");
            if (File.Exists(responseFile))
            {
                _logger.LogInformation("UserClickedOk: OK response already present for job {JobId}; allowing job to proceed.", jobId);
                _observedJobs.TryAdd(jobId, true);
                _ = _processor.EnqueueAsync(async () => await TryDeleteFileWithRetriesAsync(responseFile).ConfigureAwait(false));
                return;
            }

            try
            {
                string path = info.Path ?? string.Empty;
                if (!string.IsNullOrEmpty(path))
                {
                    // Schedule cancellation on processor to avoid spinning up many Task.Run
                    _ = _processor.EnqueueAsync(async () =>
                    {
                        try
                        {
                            using var job = new ManagementObject(path);
                            job.Delete();
                            _logger.LogInformation("JobCancelled: Successfully cancelled print job {JobId}", jobId);
                        }
                        catch (ManagementException mex)
                        {
                            _logger.LogError(mex, "JobCancelled: ManagementException while cancelling job {JobId}", jobId);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "JobCancelled: Unexpected error while cancelling job {JobId}", jobId);
                        }
                    });
                }
                else
                {
                    _logger.LogWarning("JobCancelled: Could not determine WMI path for job {JobId}; cancellation not attempted.", jobId);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "JobCancelled: Unexpected error while scheduling cancellation for job {JobId}", jobId);
            }
        }

        private void OnWatcherStopped(object? sender, StoppedEventArgs e)
        {
            try
            {
                _logger.LogWarning("ManagementEventWatcher stopped unexpectedly. Event args: {Args}. Attempting to restart.", e);
                if (sender is ManagementEventWatcher watcher)
                {
                    try
                    {
                        watcher.Start();
                        _logger.LogInformation("ManagementEventWatcher restarted successfully.");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Failed to restart ManagementEventWatcher after stop.");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error handling watcher stopped event");
            }
        }

        public override async Task StopAsync(CancellationToken cancellationToken)
        {
            await _processor.StopAsync().ConfigureAwait(false);
            DisposeWatchers();
            _logger.LogInformation("PrintMonitorWorker stopping");
            await base.StopAsync(cancellationToken).ConfigureAwait(false);
        }

        private void DisposeWatchers()
        {
            try
            {
                if (_printJobWatcher != null)
                {
                    _printJobWatcher.EventArrived -= OnPrintJobArrived;
                    _printJobWatcher.Stopped -= OnWatcherStopped;
                    try { _printJobWatcher.Stop(); } catch { }
                    _printJobWatcher.Dispose();
                    _printJobWatcher = null;
                }

                if (_responseWatcher != null)
                {
                    _responseWatcher.Created -= OnResponseFileCreated;
                    _responseWatcher.Dispose();
                    _responseWatcher = null;
                }

                if (_simulatorWatcher != null)
                {
                    _simulatorWatcher.Created -= OnSimulatorFileCreated;
                    _simulatorWatcher.Dispose();
                    _simulatorWatcher = null;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error disposing watchers");
            }
        }
    }
}
