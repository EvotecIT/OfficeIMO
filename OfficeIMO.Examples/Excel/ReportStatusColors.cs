using System;
using System.Collections.Generic;

namespace OfficeIMO.Examples.Excel;

internal static class ReportStatusColors {
    internal static (IDictionary<string, string> FillHexMap, ISet<string> BoldSet) Create(OfficeTheme theme) => (
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["Error"] = theme.Palette.Danger.ToRgbHex(),
            ["Warning"] = theme.Palette.Warning.ToRgbHex(),
            ["Success"] = theme.Palette.Success.ToRgbHex(),
            ["Ok"] = theme.Palette.Success.ToRgbHex(),
            ["Pass"] = theme.Palette.Success.ToRgbHex()
        },
        new HashSet<string>(new[] { "Error", "Warning" }, StringComparer.OrdinalIgnoreCase)
    );
}
