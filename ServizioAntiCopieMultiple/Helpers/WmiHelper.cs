using System;
using System.Management;

namespace ServizioAntiCopieMultiple.Helpers
{
    internal static class WmiHelper
    {
        public static object? GetPropertyValueSafe(ManagementBaseObject target, string propertyName)
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
    }
}
