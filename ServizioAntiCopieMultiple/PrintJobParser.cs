using System;
using System.Collections.Generic;

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
                if (name.Contains(","))
                {
                    var parts = name.Split(',');
                    return parts[^1].Trim();
                }

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
    }
}
