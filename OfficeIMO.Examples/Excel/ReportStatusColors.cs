using System;
using System.Collections.Generic;

namespace OfficeIMO.Examples.Excel;

internal static class ReportStatusColors {
    internal static (IDictionary<string, OfficeColor> FillMap, ISet<string> BoldSet) Create(OfficeTheme theme) => (
        new Dictionary<string, OfficeColor>(StringComparer.OrdinalIgnoreCase) {
            ["Error"] = theme.Palette.Danger,
            ["Warning"] = theme.Palette.Warning,
            ["Success"] = theme.Palette.Success,
            ["Ok"] = theme.Palette.Success,
            ["Pass"] = theme.Palette.Success
        },
        new HashSet<string>(new[] { "Error", "Warning" }, StringComparer.OrdinalIgnoreCase)
    );
}
