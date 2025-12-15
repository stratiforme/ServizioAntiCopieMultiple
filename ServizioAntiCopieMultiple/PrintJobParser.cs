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

        // New helper to extract copies from a WMI ManagementBaseObject without allocations
        public static int GetCopiesFromManagementObject(ManagementBaseObject? target)
        {
            if (target == null)
                return 1;

            try
            {
                var copiesObj = target["Copies"] ?? target["TotalPages"];
                if (copiesObj != null)
                {
                    if (copiesObj is int ci)
                        return ci;

                    if (int.TryParse(copiesObj.ToString(), out var parsed))
                        return parsed;
                }
            }
            catch
            {
                // ignore and return default
            }

            return 1;
        }
    }
}
