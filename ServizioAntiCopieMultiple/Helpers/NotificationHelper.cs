using System;
using System.Diagnostics;
using System.Runtime.Versioning;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;

namespace ServizioAntiCopieMultiple.Helpers
{
    [SupportedOSPlatform("windows")]
    public static class NotificationHelper
    {
        // P/Invoke for user notifications
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBox(IntPtr hWnd, string text, string caption, int type);

        private const int MB_ICONWARNING = 0x00000030;
        private const int MB_OK = 0x00000000;
        private const int MB_TOPMOST = 0x00040000;

        public static void NotifyMultipleCopiesDetected(
            string jobId, 
            string document, 
            string owner, 
            int copies,
            ILogger? logger = null)
        {
            try
            {
                // Log to Event Log so the notification is visible to administrators
                TryLogToEventLog(jobId, document, owner, copies, logger);
                
                // Also ensure it's logged at a high level so it stands out in application logs
                logger?.LogWarning("?? MULTIPLE COPIES DETECTED - Job {JobId}: {Document} ({Copies}x) by {Owner}", jobId, document, copies, owner);

                // Send desktop notification to the user
                TrySendDesktopNotification(jobId, document, owner, copies, logger);
            }
            catch (Exception ex)
            {
                logger?.LogError(ex, "Error sending notification for job {JobId}", jobId);
            }
        }

        private static void TrySendDesktopNotification(string jobId, string document, string owner, int copies, ILogger? logger)
        {
            try
            {
                string title = "?? MULTIPLE COPIES DETECTED";
                string message = $"Job ID: {jobId}\n" +
                    $"Document: {document}\n" +
                    $"User: {owner}\n" +
                    $"Copies: {copies}\n\n" +
                    $"The print job has been cancelled to prevent waste.";

                // Show warning message box (topmost to ensure visibility even if dialog is behind other windows)
                MessageBox(IntPtr.Zero, message, title, MB_ICONWARNING | MB_OK | MB_TOPMOST);
                
                logger?.LogInformation("Desktop notification shown for job {JobId}", jobId);
            }
            catch (Exception ex)
            {
                logger?.LogDebug(ex, "Failed to show desktop notification for job {JobId}", jobId);
            }
        }

        private static void TryLogToEventLog(string jobId, string document, string owner, int copies, ILogger? logger)
        {
            try
            {
                const string logName = "Application";
                const string sourceName = "ServizioAntiCopieMultiple";

                if (!EventLog.SourceExists(sourceName))
                {
                    EventLog.CreateEventSource(sourceName, logName);
                }

                string message = $"?? MULTIPLE COPIES DETECTED\n\n" +
                    $"Job ID: {jobId}\n" +
                    $"Document: {document}\n" +
                    $"User: {owner}\n" +
                    $"Copies: {copies}\n\n" +
                    $"Action: Job is being cancelled automatically.\n" +
                    $"Time: {DateTime.Now:G}";

                using (var eventLog = new EventLog(logName))
                {
                    eventLog.Source = sourceName;
                    eventLog.WriteEntry(message, EventLogEntryType.Warning, 1000);
                }

                logger?.LogInformation("Event Log notification created for job {JobId}", jobId);
            }
            catch (Exception ex)
            {
                logger?.LogDebug(ex, "Failed to write to Event Log for job {JobId}", jobId);
            }
        }
    }
}
