#nullable enable
using System;
using System.Collections.Generic;
using System.Management;
using System.Runtime.Versioning;
using ServizioAntiCopieMultiple.Helpers;

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

                // Fallback: some drivers embed copies count in the job Name (e.g. "Printer name, 10")
                if (props.TryGetValue("Name", out var nameObj) && nameObj != null)
                {
                    var nameStr = nameObj.ToString() ?? string.Empty;
                    var m = System.Text.RegularExpressions.Regex.Match(nameStr, ",\\s*(\\d+)\\s*$");
                    if (m.Success && int.TryParse(m.Groups[1].Value, out var fromName) && fromName > 0)
                        return fromName;

                    // also try any trailing number
                    var m2 = System.Text.RegularExpressions.Regex.Match(nameStr, "(\\d+)\\s*$");
                    if (m2.Success && int.TryParse(m2.Groups[1].Value, out var fromName2) && fromName2 > 0)
                        return fromName2;
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
                var copiesObj = WmiHelper.GetPropertyValueSafe(target, "Copies");
                if (copiesObj != null)
                {
                    if (copiesObj is int ci) return ci;
                    if (int.TryParse(copiesObj.ToString(), out var parsed)) return parsed;
                }

                int totalPages = 0;
                var totalObj = WmiHelper.GetPropertyValueSafe(target, "TotalPages");
                if (totalObj != null && int.TryParse(totalObj.ToString(), out var tp))
                    totalPages = tp;

                int pagesPerDoc = 0;
                string[] pagePropCandidates = { "Pages", "NumberOfPages", "PageCount" };
                foreach (var prop in pagePropCandidates)
                {
                    var pObj = WmiHelper.GetPropertyValueSafe(target, prop);
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
                        ptObj = WmiHelper.GetPropertyValueSafe(target, pname);
                        if (ptObj != null) break;
                    }

                    if (ptObj != null)
                    {
                        string ptXml = ptObj.ToString() ?? string.Empty;
                        if (!string.IsNullOrEmpty(ptXml))
                        {
                            var ptCopies = PrintTicketUtils.TryParseCopiesFromXml(ptXml);
                            if (ptCopies.HasValue && ptCopies.Value > 0) return ptCopies.Value;
                        }
                    }
                }
                catch
                {
                }

                // Fallback: parse copies from the Name property when present (many drivers append ", <copies>" to the name)
                try
                {
                    var nameObj = WmiHelper.GetPropertyValueSafe(target, "Name");
                    if (nameObj != null)
                    {
                        var nameStr = nameObj.ToString() ?? string.Empty;
                        var m = System.Text.RegularExpressions.Regex.Match(nameStr, ",\\s*(\\d+)\\s*$");
                        if (m.Success && int.TryParse(m.Groups[1].Value, out var nameCopies) && nameCopies > 0)
                            return nameCopies;

                        // also allow any trailing number
                        var m2 = System.Text.RegularExpressions.Regex.Match(nameStr, "(\\d+)\\s*$");
                        if (m2.Success && int.TryParse(m2.Groups[1].Value, out var nameCopies2) && nameCopies2 > 0)
                            return nameCopies2;
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
    }
}
