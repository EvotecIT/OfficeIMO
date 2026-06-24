namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes row-level layout and style changes that should be applied together.
    /// </summary>
    public sealed class ExcelRowLayoutOptions {
        /// <summary>
        /// Explicit row height in points. Use <see cref="ClearHeight"/> to remove a custom height.
        /// </summary>
        public double? Height { get; set; }

        /// <summary>
        /// Clears any custom row height before optional auto-fit is applied.
        /// </summary>
        public bool ClearHeight { get; set; }

        /// <summary>
        /// Auto-fits the row height after style changes are applied.
        /// </summary>
        public bool AutoFit { get; set; }

        /// <summary>
        /// Optional hidden state for the row.
        /// </summary>
        public bool? Hidden { get; set; }

        /// <summary>
        /// Optional bold font state applied to cells in the target column span.
        /// </summary>
        public bool? Bold { get; set; }

        /// <summary>
        /// Optional italic font state applied to cells in the target column span.
        /// </summary>
        public bool? Italic { get; set; }

        /// <summary>
        /// Optional underline font state applied to cells in the target column span.
        /// </summary>
        public bool? Underline { get; set; }

        /// <summary>
        /// Optional wrap-text state applied to cells in the target column span.
        /// </summary>
        public bool? WrapText { get; set; }

        /// <summary>
        /// Optional font family applied to cells in the target column span.
        /// </summary>
        public string? FontName { get; set; }

        /// <summary>
        /// Optional background color applied to cells in the target column span. Accepts RGB or ARGB hex text.
        /// </summary>
        public string? BackgroundColor { get; set; }

        /// <summary>
        /// First 1-based column affected by cell style options. Defaults to the worksheet used range.
        /// </summary>
        public int? FirstColumn { get; set; }

        /// <summary>
        /// Last 1-based column affected by cell style options. Defaults to the worksheet used range.
        /// </summary>
        public int? LastColumn { get; set; }
    }
}
