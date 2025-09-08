namespace OfficeIMO.Excel.Fluent.Report {
    /// <summary>
    /// Simple theming for report blocks. Values are conservative defaults.
    /// </summary>
    public sealed class ReportTheme {
        public string TitleColorHex { get; set; } = "#1F497D"; // dark blue
        public string SubtitleColorHex { get; set; } = "#7F7F7F"; // grey
        public string SectionHeaderFillHex { get; set; } = "#D9E1F2"; // light blue fill
        public string KeyFillHex { get; set; } = "#F2F2F2"; // light grey for key cells
        public string WarningFillHex { get; set; } = "#FFF4CE"; // light yellow
        public string ErrorFillHex   { get; set; } = "#FDE7E9"; // light red
        public string PositiveFillHex{ get; set; } = "#E7F4E4"; // light green

        public int DefaultLeftMarginColumns { get; set; } = 1; // 1-based columns
        public int DefaultContentWidthColumns { get; set; } = 10;
        public int DefaultSpacingRows { get; set; } = 1;

        public static ReportTheme Default { get; } = new ReportTheme();
    }
}
