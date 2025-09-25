namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Neutral theming type for worksheet composition. Internally maps to ReportTheme for now.
    /// </summary>
    public sealed class SheetTheme {
        /// <summary>Title font color (hex).</summary>
        public string TitleColorHex { get; set; } = "#1F497D";
        /// <summary>Subtitle font color (hex).</summary>
        public string SubtitleColorHex { get; set; } = "#7F7F7F";
        /// <summary>Fill color for section headers (hex).</summary>
        public string SectionHeaderFillHex { get; set; } = "#D9E1F2";
        /// <summary>Fill color for key cells (hex).</summary>
        public string KeyFillHex { get; set; } = "#F2F2F2";
        /// <summary>Fill color for warnings (hex).</summary>
        public string WarningFillHex { get; set; } = "#FFF4CE";
        /// <summary>Fill color for errors (hex).</summary>
        public string ErrorFillHex { get; set; } = "#FDE7E9";
        /// <summary>Fill color for positive accents (hex).</summary>
        public string PositiveFillHex { get; set; } = "#E7F4E4";

        /// <summary>Default left margin in columns.</summary>
        public int DefaultLeftMarginColumns { get; set; } = 1;
        /// <summary>Default content width in columns.</summary>
        public int DefaultContentWidthColumns { get; set; } = 10;
        /// <summary>Default spacing between blocks, in rows.</summary>
        public int DefaultSpacingRows { get; set; } = 1;

        /// <summary>Built-in default theme.</summary>
        public static SheetTheme Default { get; } = new SheetTheme();
    }
}
