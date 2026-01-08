#nullable enable
using System;
using System.Collections.Generic;
using System.Management;
using System.Runtime.Versioning;

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
            }

            return 1;
        }

        [SupportedOSPlatform("windows")]
        public static int GetCopiesFromManagementObject(ManagementBaseObject? target)
        {
            if (target == null)
                return 1;

            try
            {
                var copiesObj = GetPropertyValueSafe(target, "Copies");
                if (copiesObj != null)
                {
                    if (copiesObj is int ci) return ci;
                    if (int.TryParse(copiesObj.ToString(), out var parsed)) return parsed;
                }

                int totalPages = 0;
                var totalObj = GetPropertyValueSafe(target, "TotalPages");
                if (totalObj != null && int.TryParse(totalObj.ToString(), out var tp))
                    totalPages = tp;

                int pagesPerDoc = 0;
                string[] pagePropCandidates = { "Pages", "NumberOfPages", "PageCount" };
                foreach (var prop in pagePropCandidates)
                {
                    var pObj = GetPropertyValueSafe(target, prop);
                    if (pObj != null && int.TryParse(pObj.ToString(), out var pVal) && pVal > 0)
                    {
                        pagesPerDoc = pVal;
                        break;
                    }
                }

                if (pagesPerDoc > 0 && totalPages > 0)
                {
                    int inferred = Math.Max(1, (totalPages + pagesPerDoc - 1) / pagesPerDoc);
                    return inferred;
                }

                if (totalPages > 1)
                {
                    return Math.Min(totalPages, 1000);
                }

                try
                {
                    object? ptObj = null;
                    foreach (var pname in new[] { "PrintTicket", "PrintTicketXML", "PrintTicketData" })
                    {
                        ptObj = GetPropertyValueSafe(target, pname);
                        if (ptObj != null) break;
                    }

                    if (ptObj != null)
                    {
                        string ptXml = ptObj.ToString() ?? string.Empty;
                        if (!string.IsNullOrEmpty(ptXml))
                        {
                            var ptCopies = ParseCopiesFromPrintTicketXml(ptXml);
                            if (ptCopies.HasValue && ptCopies.Value > 0) return ptCopies.Value;
                        }
                    }
                }
                catch
                {
                }
            }
            catch
            {
            }

            return 1;
        }

        // Local fallback helper to safely read properties from ManagementBaseObject without depending on external helper class
        [SupportedOSPlatform("windows")]
        private static object? GetPropertyValueSafe(ManagementBaseObject? target, string propertyName)
        {
            if (target == null) return null;
            try
            {
                foreach (PropertyData p in target.Properties)
                {
                    if (string.Equals(p.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                    {
                        try { return p.Value; } catch { return null; }
                    }
                }
            }
            catch
            {
                // swallow exceptions - caller should handle null results
            }
            return null;
        }

        private static int? ParseCopiesFromPrintTicketXml(string xml)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(xml)) return null;

                string lower = xml.ToLowerInvariant();
                var tags = new[] { "jobcopies", "copycount", "copy-count", "copies" };
                foreach (var tag in tags)
                {
                    string open = "<" + tag + ">";
                    string close = "</" + tag + ">";
                    int oi = lower.IndexOf(open);
                    if (oi >= 0)
                    {
                        int ci = lower.IndexOf(close, oi + open.Length);
                        if (ci > oi)
                        {
                            string inner = lower.Substring(oi + open.Length, ci - (oi + open.Length)).Trim();
                            var m = System.Text.RegularExpressions.Regex.Match(inner, "\\d+");
                            if (m.Success && int.TryParse(m.Value, out var v)) return v;
                        }
                    }
                }

                var ma = System.Text.RegularExpressions.Regex.Match(lower, "copy(?:count|-count)?=\"?(\\d+)\"?" );
                if (ma.Success && int.TryParse(ma.Groups[1].Value, out var av)) return av;
            }
            catch
            {
            }

            return null;
        }
    }
}
