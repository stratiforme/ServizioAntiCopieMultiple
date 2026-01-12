using System;
using System.Management;
using System.Printing;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace ServizioAntiCopieMultiple
{
    internal sealed class PrintJobCanceller
    {
        /// <summary>
        /// Attempts to cancel a print job. Tries WMI Delete on the provided __PATH first,
        /// then falls back to a WMI search by JobId/Name/HostPrintQueue, and finally to System.Printing-based cancellation by parsing the printer/name.
        /// Returns true when cancellation is believed successful.
        /// </summary>
        public async Task<bool> CancelAsync(string? wmiPath, string? name, string owner, string? hostPrintQueue, ILogger logger)
        {
            return await Task.Run(() =>
            {
                try
                {
                    // 1) Try WMI Delete when __PATH available
                    if (!string.IsNullOrEmpty(wmiPath))
                    {
                        try
                        {
                            using var mo = new ManagementObject(wmiPath);
                            mo.Delete();
                            logger.LogDebug("PrintJobCanceller: WMI Delete succeeded for {Path}", wmiPath);
                            return true;
                        }
                        catch (ManagementException mex)
                        {
                            logger.LogDebug(mex, "PrintJobCanceller: WMI Delete failed for {Path}", wmiPath);
                        }
                        catch (Exception ex)
                        {
                            logger.LogDebug(ex, "PrintJobCanceller: unexpected error during WMI Delete for {Path}", wmiPath);
                        }
                    }

                    // Parse name for printer and job id early so we can reuse
                    int? parsedJobId = null;
                    string printerName = string.Empty;
                    if (!string.IsNullOrEmpty(name))
                    {
                        try
                        {
                            int comma = name.LastIndexOf(',');
                            if (comma >= 0)
                            {
                                printerName = name.Substring(0, comma).Trim();
                                var idStr = name.Substring(comma + 1).Trim();
                                if (int.TryParse(idStr, out var id)) parsedJobId = id;
                            }
                        }
                        catch { /* ignore parse errors */ }

                        logger.LogDebug("PrintJobCanceller: Parsed printerName='{Printer}', jobId={JobId} from name '{Name}'", printerName, parsedJobId, name);
                    }

                    // If HostPrintQueue provided, try to derive printerName/server
                    string hostServer = string.Empty;
                    string hostPrinter = string.Empty;
                    if (!string.IsNullOrEmpty(hostPrintQueue))
                    {
                        try
                        {
                            var h = hostPrintQueue.Trim();
                            if (h.StartsWith("\\\\", StringComparison.Ordinal))
                            {
                                var trimmed = h.TrimStart('\\'); // remove leading backslashes
                                var idx = trimmed.IndexOf('\\');
                                if (idx >= 0)
                                {
                                    hostServer = trimmed.Substring(0, idx);
                                    hostPrinter = trimmed.Substring(idx + 1);
                                }
                                else
                                {
                                    hostServer = trimmed;
                                }
                            }
                            else
                            {
                                // if no leading \\, treat whole as printer name
                                hostPrinter = h;
                            }

                            if (string.IsNullOrEmpty(printerName) && !string.IsNullOrEmpty(hostPrinter))
                            {
                                printerName = hostPrinter;
                                logger.LogDebug("PrintJobCanceller: Using HostPrintQueue-derived printerName='{Printer}'", printerName);
                            }
                        }
                        catch { }
                    }

                    // 2) Fallback WMI search: try to find Win32_PrintJob by JobId (or Name/HostPrintQueue) and delete it
                    if (parsedJobId.HasValue || !string.IsNullOrEmpty(printerName) || !string.IsNullOrEmpty(hostPrintQueue))
                    {
                        try
                        {
                            string wql;
                            if (parsedJobId.HasValue)
                            {
                                wql = $"SELECT * FROM Win32_PrintJob WHERE JobId = {parsedJobId.Value}";
                            }
                            else if (!string.IsNullOrEmpty(hostPrintQueue))
                            {
                                var esc = EscapeWqlLike(hostPrintQueue);
                                wql = $"SELECT * FROM Win32_PrintJob WHERE Name LIKE '%{esc}%'";
                            }
                            else
                            {
                                var esc = EscapeWqlLike(printerName);
                                wql = $"SELECT * FROM Win32_PrintJob WHERE Name LIKE '%{esc}%'";
                            }

                            logger.LogDebug("PrintJobCanceller: WMI search with query: {Query}", wql);

                            using var searcher = new ManagementObjectSearcher(new ObjectQuery(wql));
                            var results = searcher.Get();
                            foreach (ManagementObject mo in results)
                            {
                                try
                                {
                                    // Optional additional checks: Owner, Document
                                    var moOwner = mo.Properties["Owner"]?.Value?.ToString() ?? string.Empty;
                                    var moName = mo.Properties["Name"]?.Value?.ToString() ?? string.Empty;
                                    logger.LogDebug("PrintJobCanceller: WMI found object Name={Name}, Owner={Owner}", moName, moOwner);

                                    mo.Delete();
                                    logger.LogInformation("PrintJobCanceller: WMI Delete succeeded for discovered object (JobId={JobId}, Name={Name})", parsedJobId, moName);
                                    return true;
                                }
                                catch (ManagementException mex)
                                {
                                    logger.LogDebug(mex, "PrintJobCanceller: WMI Delete failed for discovered object");
                                }
                                catch (Exception ex)
                                {
                                    logger.LogDebug(ex, "PrintJobCanceller: unexpected error deleting discovered WMI object");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.LogDebug(ex, "PrintJobCanceller: WMI search fallback failed");
                        }
                    }

                    // 3) Fallback: try using System.Printing by parsing printer and job id from name
                    if (!string.IsNullOrEmpty(name) || !string.IsNullOrEmpty(hostPrinter) || !string.IsNullOrEmpty(hostServer))
                    {
                        try
                        {
                            // If hostServer is present and hostPrinter present, try remote PrintServer first
                            if (!string.IsNullOrEmpty(hostServer) && !string.IsNullOrEmpty(hostPrinter))
                            {
                                try
                                {
                                    logger.LogDebug("PrintJobCanceller: Attempting remote PrintServer {Server} for printer {Printer}", hostServer, hostPrinter);
                                    using var remote = new PrintServer($"\\\\{hostServer}");
                                    using var queue = remote.GetPrintQueue(hostPrinter);
                                    queue.Refresh();
                                    var jobs = queue.GetPrintJobInfoCollection();
                                    foreach (PrintSystemJobInfo job in jobs)
                                    {
                                        logger.LogDebug("PrintJobCanceller: Checking job {JobId} on remote queue {Printer}", job.JobIdentifier, hostPrinter);
                                        if (parsedJobId.HasValue && job.JobIdentifier == parsedJobId.Value)
                                        {
                                            try
                                            {
                                                job.Cancel();
                                                logger.LogInformation("PrintJobCanceller: Successfully cancelled job {JobId} on remote printer {Printer}", parsedJobId.Value, hostPrinter);
                                                return true;
                                            }
                                            catch (Exception jex)
                                            {
                                                logger.LogDebug(jex, "PrintJobCanceller: Cancel() failed for job {JobId} on remote printer {Printer}", parsedJobId.Value, hostPrinter);
                                            }
                                        }
                                    }
                                }
                                catch (Exception rex)
                                {
                                    logger.LogDebug(rex, "PrintJobCanceller: Remote PrintServer attempt failed for {Server}\\{Printer}", hostServer, hostPrinter);
                                }
                            }

                            using var server = new LocalPrintServer();

                            // If we have a printer name, try that queue first
                            if (!string.IsNullOrEmpty(printerName))
                            {
                                try
                                {
                                    logger.LogDebug("PrintJobCanceller: Attempting to access queue {Printer}", printerName);
                                    using var queue = server.GetPrintQueue(printerName);
                                    queue.Refresh();
                                    var jobs = queue.GetPrintJobInfoCollection();

                                    foreach (PrintSystemJobInfo job in jobs)
                                    {
                                        logger.LogDebug("PrintJobCanceller: Checking job {JobId} on queue {Printer}", job.JobIdentifier, printerName);
                                        if (parsedJobId.HasValue && job.JobIdentifier == parsedJobId.Value)
                                        {
                                            try
                                            {
                                                job.Cancel();
                                                logger.LogInformation("PrintJobCanceller: Successfully cancelled job {JobId} on printer {Printer}", parsedJobId.Value, printerName);
                                                return true;
                                            }
                                            catch (Exception jex)
                                            {
                                                logger.LogDebug(jex, "PrintJobCanceller: Cancel() failed for job {JobId} on printer {Printer}", parsedJobId.Value, printerName);
                                            }
                                        }
                                    }
                                }
                                catch (Exception qex)
                                {
                                    logger.LogDebug(qex, "PrintJobCanceller: Failed accessing queue {Printer}", printerName);
                                }
                            }

                            // As last resort, scan all queues for the job id
                            if (parsedJobId.HasValue)
                            {
                                logger.LogDebug("PrintJobCanceller: Scanning all print queues for job {JobId}", parsedJobId.Value);
                                foreach (var pq in server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections, EnumeratedPrintQueueTypes.Shared }))
                                {
                                    try
                                    {
                                        using var q = pq;
                                        q.Refresh();
                                        var jobs = q.GetPrintJobInfoCollection();
                                        foreach (PrintSystemJobInfo job in jobs)
                                        {
                                            if (job.JobIdentifier == parsedJobId.Value)
                                            {
                                                try
                                                {
                                                    job.Cancel();
                                                    logger.LogInformation("PrintJobCanceller: Successfully cancelled job {JobId} on queue {Queue}", parsedJobId.Value, q.Name);
                                                    return true;
                                                }
                                                catch (Exception jex)
                                                {
                                                    logger.LogDebug(jex, "PrintJobCanceller: Cancel() failed for job {JobId} on queue {Queue}", parsedJobId.Value, q.Name);
                                                }
                                            }
                                        }
                                    }
                                    catch { /* ignore individual queue errors */ }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.LogDebug(ex, "PrintJobCanceller: System.Printing fallback failed");
                        }
                    }

                    logger.LogDebug("PrintJobCanceller: cancellation not performed (Path='{Path}', Name='{Name}')", wmiPath, name);
                    return false;
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "PrintJobCanceller: unexpected error while attempting cancellation");
                    return false;
                }
            }).ConfigureAwait(false);
        }

        private static string EscapeWqlLike(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            // Escape single quote by doubling it for WQL string literal
            return input.Replace("'", "''");
        }
    }
}
