#nullable enable
using System;
using System.Collections.Generic;
using System.Management;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;
using ServizioAntiCopieMultiple.Helpers;

namespace ServizioAntiCopieMultiple
{
    public static class PrintJobParser
    {
        private static readonly Regex _copiesRegex = new Regex(@",\s*(\d+)\s*$", RegexOptions.Compiled);
        private static readonly Regex _trailingNumberRegex = new Regex(@"(\d+)\s*$", RegexOptions.Compiled);
        private static readonly Regex _xCopiesRegex = new Regex(@"(\d+)\s*[x×]\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex _parenCopiesRegex = new Regex(@"\((\d+)\s*(copie|copies|copy)\)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

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
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing job ID: {ex}");
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
                    if (copiesObj is int i)
                        return i;
                    if (int.TryParse(copiesObj.ToString(), out var parsed))
                        return parsed;
                }

                if (props.TryGetValue("TotalPages", out var pagesObj) && pagesObj != null)
                {
                    if (pagesObj is int ip)
                        return ip;
                    if (int.TryParse(pagesObj.ToString(), out var parsedPages))
                        return parsedPages;
                }

                if (props.TryGetValue("Name", out var nameObj) && nameObj != null)
                {
                    var nameStr = nameObj.ToString() ?? string.Empty;
                    int? nameResult = ParseCopiesFromString(nameStr);
                    if (nameResult.HasValue)
                        return nameResult.Value;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting copies from dictionary: {ex}");
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
                System.Diagnostics.Debug.WriteLine($"[ParseDebug] Copies property: {copiesObj}");
                
                if (copiesObj != null)
                {
                    if (copiesObj is int ci)
                    {
                        // Sanity: if Copies equals parsed JobId, skip (driver bug)
                        if (TryGetJobIdFromTarget(target, out var jid) && ci == jid)
                        {
                            System.Diagnostics.Debug.WriteLine($"Ignoring Copies value equal to JobId ({ci}) - possible driver bug");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[ParseDebug] Returning Copies from property: {ci}");
                            return ci;
                        }
                    }
                    else if (int.TryParse(copiesObj.ToString(), out var parsed))
                    {
                        if (TryGetJobIdFromTarget(target, out var jid) && parsed == jid)
                        {
                            System.Diagnostics.Debug.WriteLine($"Ignoring Copies value equal to JobId ({parsed}) - possible driver bug");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[ParseDebug] Returning parsed Copies: {parsed}");
                            return parsed;
                        }
                    }
                }

                int totalPages = GetIntPropertyValue(target, "TotalPages");
                int pagesPerDoc = GetPagesPerDoc(target);

                System.Diagnostics.Debug.WriteLine($"[ParseDebug] TotalPages={totalPages}, PagesPerDoc={pagesPerDoc}");

                if (pagesPerDoc > 0 && totalPages > 0)
                {
                    int inferred = Math.Max(1, (totalPages + pagesPerDoc - 1) / pagesPerDoc);
                    System.Diagnostics.Debug.WriteLine($"[ParseDebug] Inferred copies from pages: {inferred}");
                    return inferred;
                }

                if (totalPages > 1)
                {
                    int result = Math.Min(totalPages, 1000);
                    System.Diagnostics.Debug.WriteLine($"[ParseDebug] Using TotalPages as copies: {result}");
                    return result;
                }

                int? ptCopies = TryParsePrintTicket(target);
                System.Diagnostics.Debug.WriteLine($"[ParseDebug] PrintTicket copies: {(ptCopies.HasValue ? ptCopies.Value : "null")}");
                if (ptCopies.HasValue && ptCopies.Value > 0)
                    return ptCopies.Value;

                // Do NOT use the Name property fallback in case it is the printer name with job id ("Printer, 123").
                // Only use Name parsing when it contains explicit copy indicators ("3x", "(3 copie)", "copies").
                int? nameCopies = TryParseNameProperty(target);
                System.Diagnostics.Debug.WriteLine($"[ParseDebug] Name-based copies: {(nameCopies.HasValue ? nameCopies.Value : "null")}");
                if (nameCopies.HasValue)
                    return nameCopies.Value;

                System.Diagnostics.Debug.WriteLine($"[ParseDebug] No copies found, returning 1");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting copies from management object: {ex}");
            }

            return 1;
        }

        private static bool TryGetJobIdFromTarget(ManagementBaseObject? target, out int jobId)
        {
            jobId = 0;
            try
            {
                if (target == null)
                    return false;

                var name = WmiHelper.GetPropertyValueSafe(target, "Name")?.ToString();
                if (string.IsNullOrEmpty(name)) return false;
                var parsed = ParseJobId(name);
                if (int.TryParse(parsed, out var j))
                {
                    jobId = j;
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static int GetIntPropertyValue(ManagementBaseObject? target, string propertyName)
        {
            if (target == null)
                return 0;

            try
            {
                var obj = WmiHelper.GetPropertyValueSafe(target, propertyName);
                if (obj != null && int.TryParse(obj.ToString(), out var value))
                    return value;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting property {propertyName}: {ex}");
            }

            return 0;
        }

        private static int GetPagesPerDoc(ManagementBaseObject? target)
        {
            if (target == null)
                return 0;

            string[] pagePropCandidates = { "Pages", "NumberOfPages", "PageCount" };
            foreach (var prop in pagePropCandidates)
            {
                try
                {
                    var pObj = WmiHelper.GetPropertyValueSafe(target, prop);
                    if (pObj != null && int.TryParse(pObj.ToString(), out var pVal) && pVal > 0)
                        return pVal;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error getting page property {prop}: {ex}");
                }
            }

            return 0;
        }

        private static int? TryParsePrintTicket(ManagementBaseObject? target)
        {
            if (target == null)
                return null;

            try
            {
                foreach (var pname in new[] { "PrintTicket", "PrintTicketXML", "PrintTicketData" })
                {
                    var ptObj = WmiHelper.GetPropertyValueSafe(target, pname);
                    if (ptObj == null)
                        continue;

                    string ptXml = ptObj.ToString() ?? string.Empty;
                    if (!string.IsNullOrEmpty(ptXml))
                    {
                        var ptCopies = PrintTicketUtils.TryParseCopiesFromXml(ptXml);
                        if (ptCopies.HasValue && ptCopies.Value > 0)
                            return ptCopies.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing PrintTicket: {ex}");
            }

            return null;
        }

        private static int? TryParseNameProperty(ManagementBaseObject? target)
        {
            if (target == null)
                return null;

            try
            {
                var nameObj = WmiHelper.GetPropertyValueSafe(target, "Name");
                if (nameObj == null)
                    return null;

                var nameStr = nameObj.ToString() ?? string.Empty;
                return ParseCopiesFromString(nameStr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error parsing name property: {ex}");
            }

            return null;
        }

        private static int? ParseCopiesFromString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return null;

            // 1) Look for explicit formats like "3x" or "3 x" (common UI shorthand)
            var mX = _xCopiesRegex.Match(input);
            if (mX.Success && int.TryParse(mX.Groups[1].Value, out var xVal) && xVal > 0)
                return xVal;

            // 2) Look for parenthesis forms like "(3 copie)" or "(3 copies)"
            var mParen = _parenCopiesRegex.Match(input);
            if (mParen.Success && int.TryParse(mParen.Groups[1].Value, out var pVal) && pVal > 0)
                return pVal;

            // 3) Only if the string contains copy-related keywords, consider trailing-number heuristics
            if (input.IndexOf("cop", StringComparison.OrdinalIgnoreCase) >= 0 || input.IndexOf("copy", StringComparison.OrdinalIgnoreCase) >= 0 || input.IndexOf("x", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                var m = _copiesRegex.Match(input);
                if (m.Success && int.TryParse(m.Groups[1].Value, out var fromName) && fromName > 0)
                    return fromName;

                var m2 = _trailingNumberRegex.Match(input);
                if (m2.Success && int.TryParse(m2.Groups[1].Value, out var fromName2) && fromName2 > 0)
                    return fromName2;
            }

            return null;
        }
    }
}
