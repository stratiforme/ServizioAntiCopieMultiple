using System;
using System.IO;
using System.Xml;

namespace ServizioAntiCopieMultiple.Helpers
{
    internal static class PrintTicketUtils
    {
        // Lightweight helper: given a PrintTicket XML string, try to extract JobCopies or CopyCount.
        public static int? TryParseCopiesFromXml(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml)) return null;
            try
            {
                using var sr = new StringReader(xml);
                using var xr = XmlReader.Create(sr, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        var name = xr.LocalName?.ToLowerInvariant() ?? string.Empty;
                        if (name.Contains("copy") || name.Contains("jobcopies") || name.Contains("copycount"))
                        {
                            var inner = xr.ReadInnerXml();
                            if (int.TryParse(System.Text.RegularExpressions.Regex.Match(inner, "\\d+").Value, out var v))
                                return v;
                        }

                        if (xr.HasAttributes)
                        {
                            while (xr.MoveToNextAttribute())
                            {
                                var an = xr.Name.ToLowerInvariant();
                                if (an.Contains("copy"))
                                {
                                    if (int.TryParse(xr.Value, out var av)) return av;
                                }
                            }
                            xr.MoveToElement();
                        }
                    }
                }
            }
            catch
            {
                // ignore parse errors
            }

            return null;
        }
    }
}
