using System.Collections.Generic;

namespace OfficeIMO.Excel;

/// <summary>
/// Shared Excel built-in number format lookup used by style inspection, auto-fit, pivots, and image export.
/// </summary>
internal static class ExcelBuiltInNumberFormats {
    internal static readonly IReadOnlyDictionary<uint, string> Codes = new Dictionary<uint, string> {
        [0U] = "General",
        [1U] = "0",
        [2U] = "0.00",
        [3U] = "#,##0",
        [4U] = "#,##0.00",
        [9U] = "0%",
        [10U] = "0.00%",
        [11U] = "0.00E+00",
        [12U] = "# ?/?",
        [13U] = "# ??/??",
        [14U] = "m/d/yyyy",
        [15U] = "d-mmm-yy",
        [16U] = "d-mmm",
        [17U] = "mmm-yy",
        [18U] = "h:mm AM/PM",
        [19U] = "h:mm:ss AM/PM",
        [20U] = "h:mm",
        [21U] = "h:mm:ss",
        [22U] = "m/d/yyyy h:mm",
        [27U] = "yyyy/m/d",
        [30U] = "m/d/yy",
        [36U] = "m/d/yy",
        [37U] = "#,##0;(#,##0)",
        [38U] = "#,##0;[Red](#,##0)",
        [39U] = "#,##0.00;(#,##0.00)",
        [40U] = "#,##0.00;[Red](#,##0.00)",
        [45U] = "mm:ss",
        [46U] = "[h]:mm:ss",
        [47U] = "mm:ss.0",
        [48U] = "##0.0E+0",
        [49U] = "@"
    };

    internal static string? GetCode(uint numberFormatId) =>
        Codes.TryGetValue(numberFormatId, out string? code) ? code : null;

    internal static bool IsDate(uint numberFormatId) =>
        numberFormatId is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22
            or 27 or 30 or 36 or 45 or 46 or 47;

    internal static bool IsDateSystemShift(uint numberFormatId) =>
        numberFormatId is 14 or 15 or 16 or 17 or 22 or 27 or 30 or 36;
}
