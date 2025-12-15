using System;
using System.Collections.Concurrent;
using System.IO;
using System.Management;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System.Runtime.Versioning;

namespace ServizioAntiCopieMultiple
{
    [SupportedOSPlatform("windows")]
    public class PrintMonitorWorker : BackgroundService
    {
        private readonly ILogger<PrintMonitorWorker> _logger;
        private ManagementEventWatcher? _printJobWatcher;
        private FileSystemWatcher? _responseWatcher;
        private readonly ConcurrentDictionary<string, bool> _observedJobs = new();
        private readonly string _responsesDir;

        public PrintMonitorWorker(ILogger<PrintMonitorWorker> logger)
        {
            _logger = logger;
            _responsesDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "ServizioAntiCopieMultiple", "responses");
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
                // Ensure responses directory exists for user click simulation
                Directory.CreateDirectory(_responsesDir);

                // FileSystemWatcher to detect user OK responses (external UI/tool can drop files named <jobId>.ok)
                _responseWatcher = new FileSystemWatcher(_responsesDir, "*.ok")
                {
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
                };
                _responseWatcher.Created += OnResponseFileCreated;
                // Enable raising events after handlers attached to avoid race condition
                _responseWatcher.EnableRaisingEvents = true;

                // WMI event watcher for new print jobs
                string query = "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_PrintJob'";
                _printJobWatcher = new ManagementEventWatcher(new WqlEventQuery(query));
                _printJobWatcher.EventArrived += OnPrintJobArrived;
                _printJobWatcher.Start();

                _logger.LogInformation("PrintMonitorWorker started and monitoring print jobs. Responses dir: {dir}", _responsesDir);

                while (!stoppingToken.IsCancellationRequested)
                {
                    await Task.Delay(1000, stoppingToken);
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
                DisposeWatchers();
            }
        }

        private void OnResponseFileCreated(object? sender, FileSystemEventArgs e)
        {
            try
            {
                // Filename without extension expected to be job id or unique token
                string jobId = Path.GetFileNameWithoutExtension(e.Name)!;
                _logger.LogInformation("UserClickedOk: received OK response for job {JobId} (file: {FileName})", jobId, e.FullPath);

                // Mark observed so other logic can react if needed (only add once)
                _observedJobs.TryAdd(jobId, true);

                // Attempt to remove the response file to avoid accumulation. Non-critical.
                TryDeleteFileWithRetries(e.FullPath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error handling response file {FileName}", e.FullPath);
            }
        }

        private static void TryDeleteFileWithRetries(string path, int retries = 3)
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
                    // Wait briefly before retrying
                    Thread.Sleep(100);
                }
            }
        }

        private void OnPrintJobArrived(object? sender, EventArrivedEventArgs e)
        {
            try
            {
                var target = (ManagementBaseObject?)e.NewEvent?["TargetInstance"];
                if (target == null)
                    return;

                // Extract useful properties
                string? name = target["Name"]?.ToString(); // often "PrinterName, JobId"
                string document = target["Document"]?.ToString() ?? string.Empty;
                string owner = target["Owner"]?.ToString() ?? string.Empty;

                // Use shared parser helpers
                string jobId = PrintJobParser.ParseJobId(name);
                int copies = PrintJobParser.GetCopiesFromManagementObject(target);

                if (copies > 1)
                {
                    _logger.LogInformation("DetectedMultiCopyPrintJob: JobId={JobId}, Document={Document}, Owner={Owner}, Copies={Copies}", jobId, document, owner, copies);

                    _logger.LogInformation("NotificationSent: Notification sent to user for job {JobId}", jobId);

                    // Prefer in-memory observed set first to avoid a File I/O hit
                    if (_observedJobs.ContainsKey(jobId))
                    {
                        _logger.LogInformation("UserClickedOk: OK response already observed for job {JobId}; allowing job to proceed.", jobId);
                        return;
                    }

                    // If an OK response file exists immediately, respect it and do not cancel
                    string responseFile = Path.Combine(_responsesDir, jobId + ".ok");
                    if (File.Exists(responseFile))
                    {
                        _logger.LogInformation("UserClickedOk: OK response already present for job {JobId}; allowing job to proceed.", jobId);
                        _observedJobs.TryAdd(jobId, true);
                        // remove the file to keep things clean
                        TryDeleteFileWithRetries(responseFile);
                        return;
                    }

                    // Attempt to cancel the job — do not block the WMI event thread, run in background
                    try
                    {
                        string path = target["__PATH"]?.ToString() ?? string.Empty;
                        if (!string.IsNullOrEmpty(path))
                        {
                            Task.Run(() =>
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
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing print job event");
            }
        }

        public override Task StopAsync(CancellationToken cancellationToken)
        {
            DisposeWatchers();
            _logger.LogInformation("PrintMonitorWorker stopping");
            return base.StopAsync(cancellationToken);
        }

        private void DisposeWatchers()
        {
            try
            {
                if (_printJobWatcher != null)
                {
                    _printJobWatcher.EventArrived -= OnPrintJobArrived;
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
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error disposing watchers");
            }
        }
    }
}
