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
        /// Tenta di cancellare un job di stampa. Prima prova con System.Printing, se non riesce usa fallback WMI (se fornito).
        /// Ritorna true se la cancellazione è andata a buon fine.
        /// </summary>
        public async Task<bool> CancelAsync(string? wmiPath, string? name, ILogger logger)
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
                var res = await TryCancelWithWmiAsync(wmiPath, logger).ConfigureAwait(false);
                if (res)
                {
                    logger.LogDebug("PrintJobCanceller: cancellation via WMI succeeded for path {Path}", wmiPath ?? "<null>");
                    return true;
                }

                logger.LogDebug("PrintJobCanceller: WMI fallback did not cancel job for path {Path}", wmiPath ?? "<null>");
            }
            catch (Exception ex)
            {
                logger.LogDebug(ex, "PrintJobCanceller: WMI fallback threw for path {Path}", wmiPath ?? "<null>");
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

                            // Optionally match by submitter/owner if provided in job.Name string
                            // (non-intrusive: only if other matches fail)
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

        private Task<bool> TryCancelWithWmiAsync(string? wmiPath, ILogger logger)
        {
            return Task.Run(() =>
            {
                if (string.IsNullOrEmpty(wmiPath))
                    return false;

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

                return false;
            });
        }
    }
}