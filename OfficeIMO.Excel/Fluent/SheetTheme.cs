namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Neutral theming type for worksheet composition. Internally maps to ReportTheme for now.
    /// </summary>
    public sealed class SheetTheme
    {
        public string TitleColorHex { get; set; } = "#1F497D";
        public string SubtitleColorHex { get; set; } = "#7F7F7F";
        public string SectionHeaderFillHex { get; set; } = "#D9E1F2";
        public string KeyFillHex { get; set; } = "#F2F2F2";
        public string WarningFillHex { get; set; } = "#FFF4CE";
        public string ErrorFillHex   { get; set; } = "#FDE7E9";
        public string PositiveFillHex{ get; set; } = "#E7F4E4";

        public int DefaultLeftMarginColumns { get; set; } = 1;
        public int DefaultContentWidthColumns { get; set; } = 10;
        public int DefaultSpacingRows { get; set; } = 1;

        public static SheetTheme Default { get; } = new SheetTheme();
    }
}
