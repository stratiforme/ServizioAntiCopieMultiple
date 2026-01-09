ServizioAntiCopieMultiple\PrintJobCanceller.cs
using System;
using System.Linq;
using System.Management;
using System.Printing;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

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
                var res = await TryCancelWithSystemPrintingAsync(name, owner, logger).ConfigureAwait(false);
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

        private Task<bool> TryCancelWithSystemPrintingAsync(string? name, string? owner, ILogger logger)
        {
            return Task.Run(() =>
            {
                // If we don't have identifying info for System.Printing, skip quickly.
                if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(owner))
                    return false;

                try
                {
                    int comma = (name ?? string.Empty).LastIndexOf(',');
                    string printer = comma >= 0 ? name!.Substring(0, comma).Trim() : (name ?? string.Empty).Trim();
                    string idStr = comma >= 0 && comma + 1 < (name?.Length ?? 0) ? name!.Substring(comma + 1).Trim() : string.Empty;

                    using var server = new LocalPrintServer();

                    // Enumerate queues and try to match/cancel jobs
                    foreach (var q in server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections, EnumeratedPrintQueueTypes.Shared }))
                    {
                        if (q == null || string.IsNullOrEmpty(q.Name)) continue;
                        try
                        {
                            q.Refresh();
                            var jobs = q.GetPrintJobInfoCollection();
                            foreach (PrintSystemJobInfo job in jobs)
                            {
                                try
                                {
                                    if (TryMatchAndCancelJob(job, idStr, owner, printer, logger))
                                        return true;
                                }
                                catch (Exception innerEx)
                                {
                                    logger.LogDebug(innerEx, "PrintJobCanceller: unexpected error while matching/cancelling job on queue {Printer}", q.Name);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.LogDebug(ex, "PrintJobCanceller: error enumerating jobs on queue {Queue}", q.Name);
                        }
                    }
                }
                catch (PrintQueueException pqe)
                {
                    logger.LogDebug(pqe, "PrintJobCanceller: PrintQueueException in System.Printing attempt for {Name}/{Owner}", name, owner ?? "<null>");
                }
                catch (Exception ex)
                {
                    logger.LogDebug(ex, "PrintJobCanceller: unexpected exception in System.Printing attempt for {Name}/{Owner}", name, owner ?? "<null>");
                }

                return false;
            });
        }

        private static bool TryMatchAndCancelJob(PrintSystemJobInfo job, string idStr, string? owner, string? printer, ILogger logger)
        {
            try
            {
                if (!string.IsNullOrEmpty(idStr) && int.TryParse(idStr, out var parsedId))
                {
                    if (job.JobIdentifier == parsedId)
                    {
                        job.Cancel();
                        logger.LogDebug("PrintJobCanceller: cancelled by JobIdentifier on printer {Printer} (JobId={JobId})", printer ?? "<unknown>", parsedId);
                        return true;
                    }
                }

                if (!string.IsNullOrEmpty(idStr) && job.Name?.IndexOf(idStr, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    job.Cancel();
                    logger.LogDebug("PrintJobCanceller: cancelled by Name contains idStr on printer {Printer} (idStr={IdStr})", printer ?? "<unknown>", idStr);
                    return true;
                }

                if (!string.IsNullOrEmpty(owner))
                {
                    if (!string.IsNullOrEmpty(job.Submitter) && string.Equals(job.Submitter, owner, StringComparison.OrdinalIgnoreCase))
                    {
                        job.Cancel();
                        logger.LogDebug("PrintJobCanceller: cancelled by Submitter==Owner on printer {Printer} (Owner={Owner})", printer ?? "<unknown>", owner);
                        return true;
                    }

                    if (job.Name?.IndexOf(owner, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        job.Cancel();
                        logger.LogDebug("PrintJobCanceller: cancelled by Name contains Owner on printer {Printer} (Owner={Owner})", printer ?? "<unknown>", owner);
                        return true;
                    }
                }
            }
            catch (PrintJobException pje)
            {
                logger.LogDebug(pje, "PrintJobCanceller: error cancelling job (job iteration)");
            }
            catch (Exception ex)
            {
                logger.LogDebug(ex, "PrintJobCanceller: unexpected error iterating jobs (job iteration)");
            }

            return false;
        }

        private Task<bool> TryCancelWithWmiAsync(string? wmiPath, string? name, string? owner, ILogger logger)
        {
            return Task.Run(() =>
            {
                try
                {
                    var connOptions = new ConnectionOptions
                    {
                        Impersonation = ImpersonationLevel.Impersonate,
                        EnablePrivileges = true,
                        Timeout = ManagementOptions.InfiniteTimeout
                    };

                    if (!string.IsNullOrEmpty(wmiPath))
                    {
                        try
                        {
                            using var job = new ManagementObject(wmiPath);
                            try
                            {
                                if (job.Methods.Cast<MethodData>().Any(m => string.Equals(m.Name, "Cancel", StringComparison.OrdinalIgnoreCase)))
                                {
                                    try { job.InvokeMethod("Cancel", null); logger.LogDebug("PrintJobCanceller: invoked Cancel() on WMI object {Path}", wmiPath); }
                                    catch (Exception ie) { logger.LogDebug(ie, "PrintJobCanceller: Cancel() invocation failed on {Path}", wmiPath); }
                                }
                            }
                            catch { }

                            job.Delete();
                            logger.LogDebug("PrintJobCanceller: Delete() called on WMI object {Path}", wmiPath);
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
                    }

                    if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(owner))
                        return false;

                    int comma = name?.LastIndexOf(',') ?? -1;
                    string printer = comma >= 0 ? name!.Substring(0, comma).Trim() : (name ?? string.Empty).Trim();
                    string idStr = comma >= 0 && comma + 1 < (name?.Length ?? 0) ? name!.Substring(comma + 1).Trim() : string.Empty;

                    string escPrinter = (printer ?? string.Empty).Replace("'", "''");
                    string escOwner = (owner ?? string.Empty).Replace("'", "''");
                    var whereClauses = new StringBuilder();

                    if (!string.IsNullOrEmpty(idStr) && int.TryParse(idStr, out var jid))
                    {
                        whereClauses.Append($"JobId = {jid}");
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(escPrinter))
                        {
                            var lastSegment = escPrinter;
                            var idx = escPrinter.LastIndexOf('\\');
                            if (idx >= 0 && idx + 1 < escPrinter.Length)
                                lastSegment = escPrinter.Substring(idx + 1);

                            whereClauses.Append($"Name LIKE '%{lastSegment}%' OR HostPrintQueue LIKE '%{lastSegment}%'");
                        }

                        if (!string.IsNullOrEmpty(escOwner))
                        {
                            if (whereClauses.Length > 0) whereClauses.Append(" OR ");
                            whereClauses.Append($"Owner = '{escOwner}' OR Notify = '{escOwner}' OR Name LIKE '%{escOwner}%'");
                        }
                    }

                    if (whereClauses.Length == 0)
                        return false;

                    string query = $"SELECT * FROM Win32_PrintJob WHERE {whereClauses}";
                    logger.LogDebug("PrintJobCanceller: executing WMI query: {Query}", query);

                    try
                    {
                        var scope = new ManagementScope(@"\\.\root\cimv2", connOptions);
                        try { scope.Connect(); } catch { }

                        var searcher = new ManagementObjectSearcher(scope, new ObjectQuery(query));
                        foreach (ManagementObject jobObj in searcher.Get())
                        {
                            try
                            {
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
                                        if (string.IsNullOrEmpty(idStr))
                                            continue;
                                    }
                                }

                                try
                                {
                                    if (jobObj.Methods.Cast<MethodData>().Any(m => string.Equals(m.Name, "Cancel", StringComparison.OrdinalIgnoreCase)))
                                    {
                                        try { jobObj.InvokeMethod("Cancel", null); logger.LogDebug("PrintJobCanceller: invoked Cancel() on object returned by query '{Query}'", query); }
                                        catch (Exception ie) { logger.LogDebug(ie, "PrintJobCanceller: Cancel() invocation failed on object returned by query '{Query}'", query); }
                                    }
                                }
                                catch { }

                                try
                                {
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
                            finally
                            {
                                jobObj.Dispose();
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