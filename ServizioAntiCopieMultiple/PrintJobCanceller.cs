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
        /// then falls back to System.Printing-based cancellation by parsing the printer/name.
        /// Returns true when cancellation is believed successful.
        /// </summary>
        public async Task<bool> CancelAsync(string? wmiPath, string? name, string owner, ILogger logger)
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

                    // 2) Fallback: try using System.Printing by parsing printer and job id from name
                    if (!string.IsNullOrEmpty(name))
                    {
                        int? parsedJobId = null;
                        string printerName = string.Empty;

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
                        catch { /* ignore parse errors and continue to heuristics below */ }

                        // If job id still unknown, attempt to use PrintJobParser.ParseJobId (if present)
                        if (parsedJobId == null)
                        {
                            try
                            {
                                var pid = PrintJobParser.ParseJobId(name);
                                if (int.TryParse(pid, out var pidInt)) parsedJobId = pidInt;
                            }
                            catch { /* ignore */ }
                        }

                        try
                        {
                            using var server = new LocalPrintServer();

                            // If we have a printer name, try that queue first
                            if (!string.IsNullOrEmpty(printerName))
                            {
                                try
                                {
                                    using var queue = server.GetPrintQueue(printerName);
                                    queue.Refresh();
                                    var jobs = queue.GetPrintJobInfoCollection();
                                    foreach (PrintSystemJobInfo job in jobs)
                                    {
                                        if (parsedJobId.HasValue && job.JobIdentifier == parsedJobId.Value)
                                        {
                                            try
                                            {
                                                job.Cancel();
                                                logger.LogDebug("PrintJobCanceller: Cancelled job {JobId} on printer {Printer}", parsedJobId.Value, printerName);
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
                                                    logger.LogDebug("PrintJobCanceller: Cancelled job {JobId} on queue {Queue}", parsedJobId.Value, q.Name);
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
    }
}