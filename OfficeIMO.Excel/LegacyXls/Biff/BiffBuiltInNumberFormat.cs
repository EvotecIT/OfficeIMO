namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffBuiltInNumberFormat {
        private static readonly IReadOnlyDictionary<ushort, string> Codes = new Dictionary<ushort, string> {
            [0] = "General",
            [1] = "0",
            [2] = "0.00",
            [3] = "#,##0",
            [4] = "#,##0.00",
            [9] = "0%",
            [10] = "0.00%",
            [11] = "0.00E+00",
            [12] = "# ?/?",
            [13] = "# ??/??",
            [14] = "m/d/yy",
            [15] = "d-mmm-yy",
            [16] = "d-mmm",
            [17] = "mmm-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "m/d/yy h:mm",
            [37] = "#,##0 ;(#,##0)",
            [38] = "#,##0 ;[Red](#,##0)",
            [39] = "#,##0.00;(#,##0.00)",
            [40] = "#,##0.00;[Red](#,##0.00)",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [48] = "##0.0E+0",
            [49] = "@"
        };

        internal static bool TryGetCode(ushort formatId, out string? formatCode) {
            return Codes.TryGetValue(formatId, out formatCode);
        }

        internal static bool IsDateLike(ushort formatId) {
            return formatId is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22
                or 27 or 30 or 36 or 45 or 46 or 47;
        }
    }
}
