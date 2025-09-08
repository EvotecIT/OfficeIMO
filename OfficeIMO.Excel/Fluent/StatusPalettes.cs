using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Simple preset status palettes for mapping status text to fills and bold sets.
    /// </summary>
    public static class StatusPalettes
    {
        public static (IDictionary<string, string> FillHexMap, ISet<string> BoldSet) Default => (
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Error"] = "#F8C9C6",
                ["Warning"] = "#FFE59A",
                ["Success"] = "#CDEFCB",
                ["Ok"] = "#CDEFCB",
                ["Pass"] = "#CDEFCB"
            },
            new HashSet<string>(new[]{"Error","Warning"}, StringComparer.OrdinalIgnoreCase)
        );
    }
}
