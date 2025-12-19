using System;
using System.Collections.Generic;
using System.Management;

namespace ServizioAntiCopieMultiple
{
    public static class PrintJobParser
    {
        public static string ParseJobId(string? name)
        {
            if (string.IsNullOrEmpty(name))
                return Guid.NewGuid().ToString();

            try
            {
                // Avoid allocating arrays where possible: find last comma and take remainder
                int comma = name.LastIndexOf(',');
                if (comma >= 0 && comma + 1 < name.Length)
                    return name.Substring(comma + 1).Trim();

                return name;
            }
            catch
            {
                return Guid.NewGuid().ToString();
            }
        }

        public static int GetCopiesFromDictionary(Dictionary<string, object?> props)
        {
            if (props == null)
                return 1;

            try
            {
                if (props.TryGetValue("Copies", out var copiesObj) && copiesObj != null)
                {
                    if (copiesObj is int i) return i;
                    if (int.TryParse(copiesObj.ToString(), out var parsed)) return parsed;
                }

                if (props.TryGetValue("TotalPages", out var pagesObj) && pagesObj != null)
                {
                    if (pagesObj is int ip) return ip;
                    if (int.TryParse(pagesObj.ToString(), out var parsedPages)) return parsedPages;
                }
            }
            catch
            {
                // ignore and return default
            }

            return 1;
        }

        // New helper to extract copies from a WMI ManagementBaseObject with improved heuristics
        public static int GetCopiesFromManagementObject(ManagementBaseObject? target)
        {
            if (target == null)
                return 1;

            try
            {
                // 1) Prefer explicit Copies property
                var copiesObj = target["Copies"];
                if (copiesObj != null)
                {
                    if (copiesObj is int ci)
                        return ci;
                    if (int.TryParse(copiesObj.ToString(), out var parsed))
                        return parsed;
                }

                // 2) Try to infer copies from TotalPages and a per-document page count if available
                int totalPages = 0;
                var totalObj = target["TotalPages"];
                if (totalObj != null && int.TryParse(totalObj.ToString(), out var tp))
                    totalPages = tp;

                int pagesPerDoc = 0;
                string[] pagePropCandidates = { "Pages", "NumberOfPages", "PageCount" };
                foreach (var prop in pagePropCandidates)
                {
                    var pObj = target[prop];
                    if (pObj != null && int.TryParse(pObj.ToString(), out var pVal) && pVal > 0)
                    {
                        pagesPerDoc = pVal;
                        break;
                    }
                }

                if (pagesPerDoc > 0 && totalPages > 0)
                {
                    // estimate copies as totalPages / pagesPerDoc (round up)
                    int inferred = Math.Max(1, (totalPages + pagesPerDoc - 1) / pagesPerDoc);
                    return inferred;
                }

                // 3) Fallback: some drivers report TotalPages already multiplied by copies (no per-doc pages property)
                if (totalPages > 1)
                {
                    // Treat TotalPages as likely indicator of multiple copies when no other info available
                    // Cap to a reasonable max to avoid absurd values from misreported properties
                    return Math.Min(totalPages, 1000);
                }
            }
            catch
            {
                // ignore and fallthrough
            }

            return 1;
        }
    }
}
