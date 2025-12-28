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
using System.Printing;
using Microsoft.Extensions.Configuration;
using System.Runtime.InteropServices;
using System.Text;

namespace ServizioAntiCopieMultiple
{
    [SupportedOSPlatform("windows")]
    public class PrintMonitorWorker : BackgroundService
    {
        private readonly ILogger<PrintMonitorWorker> _logger;
        private ManagementEventWatcher? _printJobWatcher;
        private ManagementEventWatcher? _printJobOpWatcher;
        private FileSystemWatcher? _responseWatcher;
        private FileSystemWatcher? _simulatorWatcher;
        private readonly ConcurrentDictionary<string, bool> _observedJobs = new();
        private readonly string _responsesDir;
        private readonly string _simulatorDir;
        private readonly string _diagnosticsDir;
        private readonly PrintJobProcessor _processor = new();
        private readonly PrintJobCanceller _canceller = new();
        private readonly ConcurrentDictionary<string, ConcurrentQueue<long>> _recentJobSignatures = new();
        private readonly ConcurrentDictionary<string, JobSequence> _sequenceTrackers = new();
        private readonly TimeSpan _signatureWindow;
        private readonly int _scanIntervalSeconds;
        private readonly int _jobAgeThresholdSeconds; // ignore jobs older than this when scanning
        private readonly bool _enableScanner;

        public PrintMonitorWorker(ILogger<PrintMonitorWorker> logger, IConfiguration config)
        {
            _logger = logger;
            _responsesDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple", "responses");
            _simulatorDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple", "simulator");
            _diagnosticsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple", "diagnostics");

            // Read configuration: prefer appsettings section PrintMonitor, then environment variables, then defaults
            try
            {
                // appsetting keys: PrintMonitor:ScanIntervalSeconds, PrintMonitor:JobAgeThresholdSeconds
                int defaultScan = 5;
                int defaultAge = 30;
                int sigVal = 10;

                int scanVal = config.GetValue<int?>("PrintMonitor:ScanIntervalSeconds") ?? defaultScan;
                int ageVal = config.GetValue<int?>("PrintMonitor:JobAgeThresholdSeconds") ?? defaultAge;
                var envScan = Environment.GetEnvironmentVariable("SACM_SCAN_INTERVAL_SECONDS");
                var envAge = Environment.GetEnvironmentVariable("SACM_JOB_AGE_THRESHOLD_SECONDS");
                if (int.TryParse(envScan, out var s)) scanVal = s;
                if (int.TryParse(envAge, out var a)) ageVal = a;
                var envSig = Environment.GetEnvironmentVariable("SACM_SIGNATURE_WINDOW_SECONDS");
                if (int.TryParse(envSig, out var sv)) sigVal = sv;

                _scanIntervalSeconds = Math.Max(1, scanVal);
                _jobAgeThresholdSeconds = Math.Max(1, ageVal);
                _signatureWindow = TimeSpan.FromSeconds(Math.Max(1, sigVal));

                // Enable scanner only when interactive by default; can be forced via config/env
                bool cfgEnable = config.GetValue<bool?>("PrintMonitor:EnableScannerInService") ?? false;
                var envEnable = Environment.GetEnvironmentVariable("SACM_ENABLE_SCANNER_IN_SERVICE");
                if (bool.TryParse(envEnable, out var envBool)) cfgEnable = envBool;
                _enableScanner = cfgEnable;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to read PrintMonitor configuration; using defaults");
                _scanIntervalSeconds = 5;
                _jobAgeThresholdSeconds = 30;
                _signatureWindow = TimeSpan.FromSeconds(10);
                _enableScanner = false;
            }

            _logger.LogInformation("PrintMonitor configuration: ScanIntervalSeconds={ScanInterval}, JobAgeThresholdSeconds={JobAge}, SignatureWindowSeconds={SigWindow}", _scanIntervalSeconds, _jobAgeThresholdSeconds, _signatureWindow.TotalSeconds);
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

                // Ensure diagnostics directory exists for property dumps
                Directory.CreateDirectory(_diagnosticsDir);
                _logger.LogInformation("Diagnostics directory ensured at {dir}", _diagnosticsDir);

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

                // Start periodic queue scanner fallback only when interactive or explicitly enabled
                if (Environment.UserInteractive || _enableScanner)
                {
                    _ = Task.Run(async () => await QueueScannerLoop(stoppingToken).ConfigureAwait(false));
                    _logger.LogInformation("Queue scanner started (interactive={Interactive}, enabledByConfig={Cfg})", Environment.UserInteractive, _enableScanner);
                }
                else
                {
                    _logger.LogInformation("Queue scanner not started because running as service and not enabled by configuration. Set PrintMonitor:EnableScannerInService or SACM_ENABLE_SCANNER_IN_SERVICE=true to override.");
                }

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

        private async Task QueueScannerLoop(CancellationToken ct)
        {
            try
            {
                while (!ct.IsCancellationRequested)
                {
                    try
                    {
                        await ScanPrintQueuesAsync(ct).ConfigureAwait(false);
                    }
                    catch (OperationCanceledException) { break; }
                    catch (Exception ex)
                    {
                        _logger.LogDebug(ex, "Queue scanner encountered an error");
                    }

                    try { await Task.Delay(TimeSpan.FromSeconds(_scanIntervalSeconds), ct).ConfigureAwait(false); } catch { break; }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "QueueScannerLoop terminated unexpectedly");
            }
        }

        private Task ScanPrintQueuesAsync(CancellationToken ct)
        {
            return Task.Run(() =>
            {
                try
                {
                    using var server = new LocalPrintServer();

                    foreach (var queueInfo in server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections, EnumeratedPrintQueueTypes.Shared }))
                    {
                        if (ct.IsCancellationRequested) break;

                        // Some PrintQueueInfo entries can have null/empty Name on some platforms; skip them
                        if (string.IsNullOrEmpty(queueInfo?.Name)) continue;

                        try
                        {
                            using var queue = server.GetPrintQueue(queueInfo.Name);
                            queue.Refresh();

                            var groups = new Dictionary<string, (int Count, int JobId, string Doc, string Owner)>();

                            var jobs = queue.GetPrintJobInfoCollection();
                            foreach (PrintSystemJobInfo job in jobs)
                            {
                                try
                                {
                                    // Filter by status via reflection to avoid compile-time dependency on specific API surface
                                    try
                                    {
                                        var statusProp = job.GetType().GetProperty("JobStatus");
                                        var statusVal = statusProp?.GetValue(job);
                                        string statusStr = statusVal?.ToString() ?? string.Empty;
                                        // allow common printing/spooling statuses, ignore others
                                        if (!string.IsNullOrEmpty(statusStr) && !(statusStr.IndexOf("Print", StringComparison.OrdinalIgnoreCase) >= 0 || statusStr.IndexOf("Spool", StringComparison.OrdinalIgnoreCase) >= 0))
                                        {
                                            continue;
                                        }
                                    }
                                    catch { }

                                    // Filter by job age (if available via reflection)
                                    try
                                    {
                                        var timeProps = new[] { "TimeJobSubmitted", "TimeSubmitted", "SubmitTime", "SubmittedTime" };
                                        DateTime? submitted = null;
                                        foreach (var pname in timeProps)
                                        {
                                            var pi = job.GetType().GetProperty(pname);
                                            if (pi == null) continue;
                                            var val = pi.GetValue(job);
                                            if (val == null) continue;
                                            if (val is DateTime dt)
                                            {
                                                submitted = dt;
                                                break;
                                            }
                                            if (DateTime.TryParse(val.ToString(), out var parsed))
                                            {
                                                submitted = parsed;
                                                break;
                                            }
                                        }

                                        if (submitted.HasValue)
                                        {
                                            if ((DateTime.Now - submitted.Value).TotalSeconds > _jobAgeThresholdSeconds)
                                            {
                                                // job too old, ignore
                                                continue;
                                            }
                                        }
                                    }
                                    catch { }

                                    string doc = job.Name ?? string.Empty;
                                    string owner = job.Submitter ?? string.Empty;
                                    int jid = job.JobIdentifier;
                                    string signature = string.Join("|", queue.Name ?? string.Empty, doc, owner);

                                    if (!groups.TryGetValue(signature, out var entry))
                                    {
                                        groups[signature] = (1, jid, doc, owner);
                                    }
                                    else
                                    {
                                        groups[signature] = (entry.Count + 1, entry.JobId, entry.Doc, entry.Owner);
                                    }
                                }
                                catch (Exception jex)
                                {
                                    _logger.LogDebug(jex, "Error iterating print jobs on queue {Queue}", queueInfo.Name);
                                }
                            }

                            foreach (var kv in groups)
                            {
                                var sig = kv.Key;
                                var val = kv.Value;
                                if (val.Count > 1)
                                {
                                    try
                                    {
                                        var info = new PrintJobInfo
                                        {
                                            Name = $"{queue.Name ?? string.Empty}, {val.JobId}",
                                            Document = val.Doc,
                                            Owner = val.Owner,
                                            Copies = val.Count,
                                            Path = string.Empty
                                        };

                                        _logger.LogInformation("ScannerDetectedMultiCopy: Printer={Printer}, Doc={Document}, Owner={Owner}, Copies={Copies}", queue.Name, val.Doc, val.Owner, val.Count);

                                        // Hand off to existing processing logic
                                        ProcessPrintJobInfo(info);
                                    }
                                    catch (Exception ex)
                                    {
                                        _logger.LogError(ex, "Error processing scanner-detected multi-copy job for signature {Sig}", sig);
                                    }
                                }
                            }
                        }
                        catch (Exception qex)
                        {
                            _logger.LogDebug(qex, "Error scanning queue {Queue}", queueInfo.Name);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "ScanPrintQueuesAsync failed");
                }
            });
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

                    // Also create a broader operation watcher to catch modifications/deletions which some drivers use
                    try
                    {
                        var opQuery = new WqlEventQuery("SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'");
                        _printJobOpWatcher = new ManagementEventWatcher(scope, opQuery);
                        _printJobOpWatcher.EventArrived += OnPrintJobOperationArrived;
                        _printJobOpWatcher.Stopped += OnWatcherStopped;
                    }
                    catch (Exception opEx)
                    {
                        _logger.LogDebug(opEx, "Failed to create operation watcher (scoped)");
                    }

                    try
                    {
                        _printJobWatcher.Start();
                        _printJobOpWatcher?.Start();
                        _logger.LogInformation("ManagementEventWatcher started with query: {Query}", query);
                        if (_printJobOpWatcher != null)
                            _logger.LogInformation("ManagementEventWatcher (operation) started with query: SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'");
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

                            // fallback op watcher (unscoped)
                            try
                            {
                                var opQuery = new WqlEventQuery("SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'");
                                _printJobOpWatcher = new ManagementEventWatcher(opQuery);
                                _printJobOpWatcher.EventArrived += OnPrintJobOperationArrived;
                                _printJobOpWatcher.Stopped += OnWatcherStopped;
                            }
                            catch (Exception opEx)
                            {
                                _logger.LogDebug(opEx, "Failed to create operation watcher (fallback)");
                            }

                            _printJobWatcher.Start();
                            _printJobOpWatcher?.Start();
                            _logger.LogInformation("ManagementEventWatcher started (fallback) with query: {Query}", query);
                            if (_printJobOpWatcher != null)
                                _logger.LogInformation("ManagementEventWatcher (operation fallback) started with query: SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'");
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

                        // unscoped op watcher
                        try
                        {
                            var opQuery = new WqlEventQuery("SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'");
                            _printJobOpWatcher = new ManagementEventWatcher(opQuery);
                            _printJobOpWatcher.EventArrived += OnPrintJobOperationArrived;
                            _printJobOpWatcher.Stopped += OnWatcherStopped;
                        }
                        catch (Exception opEx)
                        {
                            _logger.LogDebug(opEx, "Failed to create operation watcher (unscoped fallback)");
                        }

                        _printJobWatcher.Start();
                        _printJobOpWatcher?.Start();
                        _logger.LogInformation("ManagementEventWatcher started (unscoped fallback) with query: {Query}", query);
                        if (_printJobOpWatcher != null)
                            _logger.LogInformation("ManagementEventWatcher (operation unscoped fallback) started with query: SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'");
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

                // Try native DEVMODE/GetJob
                if (copies <= 1)
                {
                    try
                    {
                        var native = NativeSpool.TryGetCopiesFromW32Job(name, out var nativeDebug);
                        if (!string.IsNullOrEmpty(nativeDebug) && _logger.IsEnabled(Microsoft.Extensions.Logging.LogLevel.Debug))
                        {
                            _logger.LogDebug("NativeSpool debug: {Debug}", nativeDebug);
                        }

                        if (native.HasValue && native.Value > copies)
                        {
                            _logger.LogInformation("DEVMODECopiesDetected: increased copies from {Old} to {New} for job {JobId}", copies, native.Value, jobId);
                            copies = native.Value;
                        }
                    }
                    catch (Exception ex) { _logger.LogDebug(ex, "DEVMODE detection failed for job {JobId}", jobId); }
                }

                // 2) Try PrintTicket XML
                if (copies <= 1)
                {
                    try
                    {
                        var ptObj = target["PrintTicket"] ?? target["PrintTicketXML"] ?? target["PrintTicketData"];
                        if (ptObj != null)
                        {
                            string ptXml = ptObj.ToString() ?? string.Empty;

                            // Save print ticket to diagnostics for deeper analysis
                            try
                            {
                                if (!string.IsNullOrEmpty(ptXml))
                                {
                                    string ptFile = Path.Combine(_diagnosticsDir, $"printticket_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}_{jobId}.xml");
                                    File.WriteAllText(ptFile, ptXml);
                                    _logger.LogInformation("DiagnosticsDumpSaved: PrintTicket XML saved to {Path}", ptFile);

                                    var parsed = PrintTicketUtils.TryParseCopiesFromXml(ptXml);
                                    if (parsed.HasValue)
                                    {
                                        _logger.LogDebug("PrintTicketParser: parsed copies={Copies} from PrintTicket for job {JobId}", parsed.Value, jobId);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogDebug(ex, "Failed saving PrintTicket XML for diagnostics");
                            }

                            var parsed2 = PrintTicketUtils.TryParseCopiesFromXml(ptXml);
                            if (parsed2.HasValue && parsed2.Value > copies)
                            {
                                _logger.LogInformation("PrintTicketCopiesDetected: increased copies from {Old} to {New} for job {JobId}", copies, parsed2.Value, jobId);
                                copies = parsed2.Value;
                            }
                        }
                    }
                    catch (Exception ex) { _logger.LogDebug(ex, "PrintTicket detection failed for job {JobId}", jobId); }
                }

                // 3) Aggregate recent events (signature heuristic)
                try
                {
                    string printer = !string.IsNullOrEmpty(name) ? (name.LastIndexOf(',') >= 0 ? name.Substring(0, name.LastIndexOf(',')).Trim() : name) : string.Empty;
                    string signature = string.Join("|", printer, document ?? string.Empty, owner ?? string.Empty);
                    long now = DateTime.UtcNow.Ticks;
                    var queue = _recentJobSignatures.GetOrAdd(signature, _ => new ConcurrentQueue<long>());
                    queue.Enqueue(now);
                    while (queue.TryPeek(out long t) && TimeSpan.FromTicks(now - t) > _signatureWindow) queue.TryDequeue(out _);
                    int recentCount = queue.Count;
                    if (recentCount > copies)
                    {
                        _logger.LogInformation("InferredCopies: increased copies from {Old} to {New} based on {Count} recent events for signature {Sig}", copies, recentCount, recentCount, signature);
                        copies = recentCount;
                    }
                }
                catch (Exception ex) { _logger.LogDebug(ex, "Failed inferring copies from recent signatures"); }

                // Log received event at Information level so events are visible in logs even when Copies == 1
                try
                {
                    string wmiPath = target["__PATH"]?.ToString() ?? string.Empty;
                    _logger.LogInformation("WMIPrintEvent: JobId={JobId}, Name={Name}, Document={Document}, Owner={Owner}, Copies={Copies}, Path={Path}", jobId, name ?? "<null>", document, owner, copies, string.IsNullOrEmpty(wmiPath) ? "<empty>" : wmiPath);
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Failed to log WMIPrintEvent basic info");
                }

                // Save full TargetInstance dump to diagnostics for offline analysis
                try
                {
                    SaveTargetDumpToFile(target, jobId);
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "Failed to save TargetInstance dump to diagnostics");
                }

                // Build PrintJobInfo and hand off to common processor
                var info = new PrintJobInfo
                {
                    Name = name, // nullable allowed
                    Document = document ?? string.Empty,
                    Owner = owner ?? string.Empty,
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

                    // Log some additional commonly useful fields at debug level for easier inspection
                    void TryLogField(string field)
                    {
                        if (props.TryGetValue(field, out var v) && !string.IsNullOrEmpty(v))
                        {
                            _logger.LogDebug("TargetInstance field: {Field} = {Value}", field, v);
                        }
                    }

                    TryLogField("PrintProcessor");
                    TryLogField("Parameters");
                    TryLogField("Notify");
                    TryLogField("SubmissionTime");
                    TryLogField("TimeSubmitted");
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

                // DIAGNOSTICA MIGLIORATA: loggare Name e Path a Info (non solo Debug)
                _logger.LogInformation("AttemptingCancel: JobId={JobId}, Name={Name}, Path={Path}", jobId, info.Name ?? "<null>", string.IsNullOrEmpty(path) ? "<empty>" : path);

                if (!string.IsNullOrEmpty(path))
                {
                    // Schedule cancellation on processor to avoid spinning up many Task.Run
                    _ = _processor.EnqueueAsync(async () =>
                    {
                        try
                        {
                            var cancelled = await _canceller.CancelAsync(path, info.Name, _logger).ConfigureAwait(false);
                            if (cancelled)
                            {
                                _logger.LogInformation("JobCancelled: Successfully cancelled print job {JobId}", jobId);
                            }
                            else
                            {
                                _logger.LogWarning("JobCancelled: Cancellation attempt failed for job {JobId}. Will not retry automatically.", jobId);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "JobCancelled: Unexpected error while cancelling job {JobId}", jobId);
                        }
                    });

                    _logger.LogInformation("AttemptingCancelScheduled: cancellation scheduled for job {JobId}", jobId);
                }
                else
                {
                    _logger.LogWarning("JobCancelled: Could not determine WMI path for job {JobId}; cancellation not attempted. Name={Name}", jobId, info.Name);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "JobCancelled: Unexpected error while scheduling cancellation for job {JobId}", jobId);
            }
        }

        private void OnPrintJobOperationArrived(object? sender, EventArrivedEventArgs e)
        {
            try
            {
                var ev = e.NewEvent;
                string? eventClass = ev?.ClassPath?.ClassName;
                _logger.LogInformation("WMIPrintOpEvent: EventClass={EventClass}, TimeGenerated={Time}", eventClass, ev?["TIME_CREATED"]);

                var target = (ManagementBaseObject?)e.NewEvent?["TargetInstance"];
                if (target == null)
                {
                    _logger.LogWarning("Print job operation event arrived but TargetInstance was null");
                    return;
                }

                // For creation and modification events, reuse existing handler logic
                if (string.Equals(eventClass, "__InstanceCreationEvent", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(eventClass, "__InstanceModificationEvent", StringComparison.OrdinalIgnoreCase))
                {
                    // Build minimal PrintJobInfo and process similarly to creation
                    string? name = target["Name"]?.ToString();
                    int copies = PrintJobParser.GetCopiesFromManagementObject(target);
                    var info = new PrintJobInfo
                    {
                        Name = name,
                        Document = target["Document"]?.ToString() ?? string.Empty,
                        Owner = target["Owner"]?.ToString() ?? string.Empty,
                        Copies = copies,
                        Path = target["__PATH"]?.ToString() ?? string.Empty
                    };

                    // Log operation-level event
                    _logger.LogInformation("WMIPrintOpEventDetailed: Event={Event}, JobId={JobId}, Name={Name}, Copies={Copies}, Path={Path}", eventClass, PrintJobParser.ParseJobId(name), name ?? "<null>", copies, info.Path ?? "<empty>");

                    // Save full TargetInstance dump for operation events as well
                    try
                    {
                        SaveTargetDumpToFile(target, PrintJobParser.ParseJobId(name));
                    }
                    catch (Exception ex)
                    {
                        _logger.LogDebug(ex, "Failed to save TargetInstance dump (operation)");
                    }

                    if (info.Copies > 1)
                    {
                        ProcessPrintJobInfo(info);
                    }
                }
                else
                {
                    // Other operation events (e.g. deletion) are useful to log for diagnostics but not processed
                    _logger.LogInformation("WMIPrintOpEventIgnored: Event={EventClass} for print job (no processing performed)", eventClass);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing print job operation event");
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

                if (_printJobOpWatcher != null)
                {
                    _printJobOpWatcher.EventArrived -= OnPrintJobOperationArrived;
                    _printJobOpWatcher.Stopped -= OnWatcherStopped;
                    try { _printJobOpWatcher.Stop(); } catch { }
                    _printJobOpWatcher.Dispose();
                    _printJobOpWatcher = null;
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

        private void SaveTargetDumpToFile(ManagementBaseObject target, string jobId)
        {
            try
            {
                var props = new Dictionary<string, object?>();
                foreach (PropertyData p in target.Properties)
                {
                    try
                    {
                        props[p.Name] = p.Value;
                    }
                    catch (Exception ex)
                    {
                        props[p.Name] = $"<error: {ex.Message}>";
                    }
                }

                string fileName = $"wmi_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}_{jobId ?? "unknown"}.json";
                string path = Path.Combine(_diagnosticsDir, fileName);
                var opts = new JsonSerializerOptions { WriteIndented = true };
                File.WriteAllText(path, JsonSerializer.Serialize(props, opts));
                _logger.LogInformation("DiagnosticsDumpSaved: WMI TargetInstance properties saved to {Path}", path);

                // If PrintTicket present, also save separately (helps offline analysis)
                try
                {
                    var ptObj = target["PrintTicket"] ?? target["PrintTicketXML"] ?? target["PrintTicketData"];
                    if (ptObj != null)
                    {
                        string ptXml = ptObj.ToString() ?? string.Empty;
                        if (!string.IsNullOrEmpty(ptXml))
                        {
                            string ptFile = Path.Combine(_diagnosticsDir, $"printticket_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}_{jobId}.xml");
                            File.WriteAllText(ptFile, ptXml);
                            _logger.LogInformation("DiagnosticsDumpSaved: PrintTicket XML saved to {Path}", ptFile);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug(ex, "SaveTargetDumpToFile: failed saving PrintTicket");
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "SaveTargetDumpToFile: failed");
            }
        }

        internal sealed class JobSequence
        {
            public int LastId = 0;
            public long LastTicks = 0L;
            public int Count = 0;
            public object Lock = new object();
        }

        // Native spool helper: best-effort read of JOB_INFO_2 / DEVMODE via GetJob
        private static class NativeSpool
        {
            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            internal struct DEVMODE
            {
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string dmDeviceName;
                public short dmSpecVersion;
                public short dmDriverVersion;
                public short dmSize;
                public short dmDriverExtra;
                public int dmFields;
                public short dmOrientation;
                public short dmPaperSize;
                public short dmPaperLength;
                public short dmPaperWidth;
                public short dmScale;
                public short dmCopies;
                // rest omitted
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            internal struct JOB_INFO_2
            {
                public IntPtr pPrinterName;
                public IntPtr pMachineName;
                public IntPtr pUserName;
                public IntPtr pDocument;
                public IntPtr pDatatype;
                public IntPtr pStatus;
                public IntPtr pStatusMask;
                public int JobId;
                public IntPtr pJobTime;
                public IntPtr pSubmitted;
                public int TotalPages;
                public int PagesPrinted;
                public IntPtr pDevMode;
                // rest omitted
            }

            [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
            internal static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);
            [DllImport("winspool.drv", SetLastError = true)]
            internal static extern bool ClosePrinter(IntPtr hPrinter);
            [DllImport("winspool.drv", SetLastError = true)]
            internal static extern bool GetJob(IntPtr hPrinter, int JobId, int Level, IntPtr pJob, int cbBuf, out int pcbNeeded);

            internal static int? TryGetCopiesFromW32Job(string? name)
            {
                return TryGetCopiesFromW32Job(name, out _);
            }

            internal static int? TryGetCopiesFromW32Job(string? name, out string debug)
            {
                var sb = new StringBuilder();
                debug = string.Empty;
                try
                {
                    if (string.IsNullOrEmpty(name))
                    {
                        sb.AppendLine("name null or empty");
                        debug = sb.ToString();
                        return null;
                    }
                    int comma = name.LastIndexOf(',');
                    string printer = comma >= 0 ? name.Substring(0, comma).Trim() : name.Trim();
                    string idStr = comma >= 0 && comma + 1 < name.Length ? name.Substring(comma + 1).Trim() : string.Empty;
                    sb.AppendLine($"Parsed printer='{printer}', idStr='{idStr}'");
                    if (string.IsNullOrEmpty(printer) || string.IsNullOrEmpty(idStr) || !int.TryParse(idStr, out var jid))
                    {
                        sb.AppendLine("Could not parse printer or job id from name");
                        debug = sb.ToString();
                        return null;
                    }

                    // Try open printer as-is
                    if (!OpenPrinter(printer, out var hPrinter, IntPtr.Zero))
                    {
                        int err = Marshal.GetLastWin32Error();
                        sb.AppendLine($"OpenPrinter failed for '{printer}' with error {err}");

                        // Try alternative with server share prefix (common when printer name contains backslashes)
                        string alt = printer;
                        if (!printer.StartsWith("\\\\") && printer.Contains("\\"))
                        {
                            // try with double-escaped prefix
                            alt = "\\\\" + printer.TrimStart('\\');
                            sb.AppendLine($"Attempting OpenPrinter with alternative '{alt}'");
                            if (!OpenPrinter(alt, out hPrinter, IntPtr.Zero))
                            {
                                int err2 = Marshal.GetLastWin32Error();
                                sb.AppendLine($"OpenPrinter alternative failed with error {err2}");
                                debug = sb.ToString();
                                return null;
                            }
                        }
                        else
                        {
                            debug = sb.ToString();
                            return null;
                        }
                    }

                    try
                    {
                        int needed = 0;
                        if (!GetJob(hPrinter, jid, 2, IntPtr.Zero, 0, out needed))
                        {
                            int err = Marshal.GetLastWin32Error();
                            sb.AppendLine($"GetJob initial call returned false, needed={needed}, err={err}");
                            if (needed <= 0)
                            {
                                debug = sb.ToString();
                                return null;
                            }
                        }

                        sb.AppendLine($"Allocating buffer of size {needed}");
                        var p = Marshal.AllocHGlobal(needed);
                        try
                        {
                            if (!GetJob(hPrinter, jid, 2, p, needed, out needed))
                            {
                                int err = Marshal.GetLastWin32Error();
                                sb.AppendLine($"GetJob second call failed err={err}");
                                debug = sb.ToString();
                                return null;
                            }

                            var ji = Marshal.PtrToStructure<JOB_INFO_2>(p);
                            sb.AppendLine($"JOB_INFO_2.JobId={ji.JobId}, TotalPages={ji.TotalPages}, PagesPrinted={ji.PagesPrinted}, pDevMode={ji.pDevMode}");

                            if (ji.pDevMode != IntPtr.Zero)
                            {
                                try
                                {
                                    var dev = Marshal.PtrToStructure<DEVMODE>(ji.pDevMode);
                                    sb.AppendLine($"DEVMODE.dmCopies={dev.dmCopies}, dmSize={dev.dmSize}, dmDriverExtra={dev.dmDriverExtra}");
                                    if (dev.dmCopies > 0)
                                    {
                                        debug = sb.ToString();
                                        return dev.dmCopies;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    sb.AppendLine($"Failed marshalling DEVMODE: {ex.Message}");
                                }
                            }

                            if (ji.TotalPages > 1)
                            {
                                sb.AppendLine($"Using JOB_INFO_2.TotalPages={ji.TotalPages} as copies");
                                debug = sb.ToString();
                                return ji.TotalPages;
                            }

                            // As fallback, try to read document name and other fields by marshalling strings
                            try
                            {
                                string doc = ji.pDocument != IntPtr.Zero ? Marshal.PtrToStringUni(ji.pDocument) ?? string.Empty : string.Empty;
                                sb.AppendLine($"Document='{doc}'");
                            }
                            catch { }
                        }
                        finally
                        {
                            Marshal.FreeHGlobal(p);
                        }
                    }
                    finally
                    {
                        ClosePrinter(hPrinter);
                    }
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"Unexpected exception: {ex.Message}");
                }

                debug = sb.ToString();
                return null;
            }
        }
    }
}
