using System;
using System.Management;
using System.Printing;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using System.Text;

namespace ServizioAntiCopieMultiple
{
    internal sealed class PrintJobCanceller
    {
        /// <summary>
        /// Tenta di cancellare un job di stampa. Prima prova con System.Printing, se non riesce usa fallback WMI (se fornito).
        /// Ritorna true se la cancellazione è andata a buon fine.
        /// </summary>
        public async Task<bool> CancelAsync(string? wmiPath, string? name, string? owner, ILogger logger)
        {
            // 1) Prova System.Printing
            try
            {
                var res = await TryCancelWithSystemPrintingAsync(name, logger).ConfigureAwait(false);
                if (res)
                {
                    logger.LogDebug("PrintJobCanceller: cancellation via System.Printing succeeded for {Name}", name ?? "<null>");
                    return true;
                }

                logger.LogDebug("PrintJobCanceller: System.Printing did not find/clear the job for {Name}", name ?? "<null>");
            }
            catch (Exception ex)
            {
                logger.LogDebug(ex, "PrintJobCanceller: System.Printing attempt threw for {Name}", name ?? "<null>");
            }

            // 2) Fallback WMI
            try
            {
                var res = await TryCancelWithWmiAsync(wmiPath, name, owner, logger).ConfigureAwait(false);
                if (res)
                {
                    logger.LogDebug("PrintJobCanceller: cancellation via WMI succeeded for path/name {PathOrName}", string.IsNullOrEmpty(wmiPath) ? name ?? "<null>" : wmiPath ?? "<null>");
                    return true;
                }

                logger.LogDebug("PrintJobCanceller: WMI fallback did not cancel job for path/name {PathOrName}", string.IsNullOrEmpty(wmiPath) ? name ?? "<null>" : wmiPath ?? "<null>");
            }
            catch (Exception ex)
            {
                logger.LogDebug(ex, "PrintJobCanceller: WMI fallback threw for path/name {PathOrName}", string.IsNullOrEmpty(wmiPath) ? name ?? "<null>" : wmiPath ?? "<null>");
            }

            return false;
        }

        private Task<bool> TryCancelWithSystemPrintingAsync(string? name, ILogger logger)
        {
            return Task.Run(() =>
            {
                if (string.IsNullOrEmpty(name))
                    return false;

                try
                {
                    // Parse printer name and job id from the WMI-style Name: "PrinterName, 123"
                    int comma = name.LastIndexOf(',');
                    string printer = comma >= 0 ? name.Substring(0, comma).Trim() : name.Trim();
                    string idStr = comma >= 0 && comma + 1 < name.Length ? name.Substring(comma + 1).Trim() : string.Empty;

                    if (string.IsNullOrEmpty(printer))
                        return false;

                    using var server = new LocalPrintServer();
                    using var queue = server.GetPrintQueue(printer);
                    queue.Refresh();

                    var jobs = queue.GetPrintJobInfoCollection();
                    foreach (PrintSystemJobInfo job in jobs)
                    {
                        try
                        {
                            // Prefer match by numeric JobIdentifier if possible
                            if (!string.IsNullOrEmpty(idStr) && int.TryParse(idStr, out var parsedId))
                            {
                                if (job.JobIdentifier == parsedId)
                                {
                                    job.Cancel();
                                    return true;
                                }
                            }

                            // Fallbacks: match by Name or Submitter/Document heuristics
                            if (!string.IsNullOrEmpty(idStr) && job.Name?.Contains(idStr, StringComparison.OrdinalIgnoreCase) == true)
                            {
                                job.Cancel();
                                return true;
                            }
                        }
                        catch (PrintJobException pje)
                        {
                            logger.LogDebug(pje, "PrintJobCanceller: error cancelling job on queue {Printer} (job iteration)", printer);
                        }
                        catch (Exception ex)
                        {
                            logger.LogDebug(ex, "PrintJobCanceller: unexpected error iterating jobs on {Printer}", printer);
                        }
                    }
                }
                catch (PrintQueueException pqe)
                {
                    logger.LogDebug(pqe, "PrintJobCanceller: PrintQueueException for printer {Name}", name);
                }
                catch (Exception ex)
                {
                    logger.LogDebug(ex, "PrintJobCanceller: unexpected exception in System.Printing attempt for {Name}", name);
                }

                return false;
            });
        }

        private Task<bool> TryCancelWithWmiAsync(string? wmiPath, string? name, string? owner, ILogger logger)
        {
            return Task.Run(() =>
            {
                try
                {
                    // 1) If a WMI path was provided, try direct delete first
                    if (!string.IsNullOrEmpty(wmiPath))
                    {
                        try
                        {
                            using var job = new ManagementObject(wmiPath);
                            job.Delete();
                            return true;
                        }
                        catch (ManagementException mex)
                        {
                            logger.LogDebug(mex, "PrintJobCanceller: ManagementException while deleting WMI object {Path}", wmiPath);
                        }
                        catch (Exception ex)
                        {
                            logger.LogDebug(ex, "PrintJobCanceller: unexpected exception while deleting WMI object {Path}", wmiPath);
                        }

                        // Fall through to attempt lookup by query if direct delete failed
                    }

                    // 2) If no wmiPath (or direct delete failed), try to find the job via query using JobId/Name/HostPrintQueue/Owner/Notify
                    if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(owner))
                        return false;

                    // Parse printer and job id from name when available
                    int comma = name?.LastIndexOf(',') ?? -1;
                    string printer = comma >= 0 ? name!.Substring(0, comma).Trim() : (name ?? string.Empty).Trim();
                    string idStr = comma >= 0 && comma + 1 < (name?.Length ?? 0) ? name!.Substring(comma + 1).Trim() : string.Empty;

                    // Prepare escaped values
                    string escPrinter = (printer ?? string.Empty).Replace("'", "''");
                    string escOwner = (owner ?? string.Empty).Replace("'", "''");
                    var whereClauses = new StringBuilder();

                    if (!string.IsNullOrEmpty(idStr) && int.TryParse(idStr, out var jid))
                    {
                        whereClauses.Append($"JobId = {jid}");
                    }
                    else
                    {
                        // Add several tolerant matches combining Name, HostPrintQueue, Owner and Notify
                        if (!string.IsNullOrEmpty(escPrinter))
                        {
                            // Also try to match last segment of printer name (strip server prefix if present)
                            var lastSegment = escPrinter;
                            var idx = escPrinter.LastIndexOf('\\');
                            if (idx >= 0 && idx + 1 < escPrinter.Length)
                                lastSegment = escPrinter.Substring(idx + 1);

                            whereClauses.Append($"Name LIKE '%{lastSegment}%' OR HostPrintQueue LIKE '%{lastSegment}%'");
                        }

                        if (!string.IsNullOrEmpty(escOwner))
                        {
                            if (whereClauses.Length > 0) whereClauses.Append(" OR ");
                            // Exact owner match or Notify field
                            whereClauses.Append($"Owner = '{escOwner}' OR Notify = '{escOwner}' OR Name LIKE '%{escOwner}%'");
                        }
                    }

                    if (whereClauses.Length == 0)
                        return false;

                    string query = $"SELECT * FROM Win32_PrintJob WHERE {whereClauses}";
                    logger.LogDebug("PrintJobCanceller: executing WMI query: {Query}", query);

                    try
                    {
                        var searcher = new ManagementObjectSearcher(@"\\.\root\cimv2", query);
                        foreach (ManagementObject jobObj in searcher.Get())
                        {
                            try
                            {
                                // Best-effort verification when possible
                                if (!string.IsNullOrEmpty(idStr) && int.TryParse(idStr, out var parsedIdCheck))
                                {
                                    var jobIdProp = jobObj["JobId"];
                                    if (jobIdProp == null || !int.TryParse(jobIdProp.ToString(), out var foundId) || foundId != parsedIdCheck)
                                        continue;
                                }

                                if (!string.IsNullOrEmpty(escOwner))
                                {
                                    var ownerProp = jobObj["Owner"]?.ToString();
                                    var notifyProp = jobObj["Notify"]?.ToString();
                                    if (!string.Equals(ownerProp, owner, StringComparison.OrdinalIgnoreCase)
                                        && !string.Equals(notifyProp, owner, StringComparison.OrdinalIgnoreCase))
                                    {
                                        // If owner was requested but doesn't match, try next candidate
                                        // however allow deletion if other indicators (JobId) matched earlier
                                        if (string.IsNullOrEmpty(idStr))
                                            continue;
                                    }
                                }

                                jobObj.Delete();
                                logger.LogDebug("PrintJobCanceller: WMI delete called on object matching query '{Query}'", query);
                                return true;
                            }
                            catch (ManagementException mexInner)
                            {
                                logger.LogDebug(mexInner, "PrintJobCanceller: ManagementException deleting object returned by query '{Query}'", query);
                            }
                            catch (Exception exInner)
                            {
                                logger.LogDebug(exInner, "PrintJobCanceller: unexpected exception deleting object returned by query '{Query}'", query);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.LogDebug(ex, "PrintJobCanceller: WMI query failed: {Query}", query);
                    }
                }
                catch (Exception ex)
                {
                    logger.LogDebug(ex, "PrintJobCanceller: unexpected exception in TryCancelWithWmiAsync fallback");
                }

                return false;
            });
        }
    }
}