namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for clamping column widths and enabling wrapping by header.
    /// Preset defaults use character units; each width also accepts a physical length.
    /// </summary>
    public sealed class ExcelColumnSizingOptions {
        /// <summary>
        /// Character or physical width used for headers registered in <see cref="ShortHeaders"/>.
        /// </summary>
        public ExcelColumnWidth ShortWidth { get; set; } = ExcelColumnWidth.Characters(16);     // Status/Alg/Hash

        /// <summary>
        /// Character or physical width used for headers registered in <see cref="NumericHeaders"/>.
        /// </summary>
        public ExcelColumnWidth NumericWidth { get; set; } = ExcelColumnWidth.Characters(10);   // Depth/Count/Days

        /// <summary>
        /// Character or physical width used for headers registered in <see cref="MediumHeaders"/>.
        /// </summary>
        public ExcelColumnWidth MediumWidth { get; set; } = ExcelColumnWidth.Characters(28);    // Provider/Target/Reason

        /// <summary>
        /// Character or physical width used for headers registered in <see cref="LongHeaders"/>.
        /// </summary>
        public ExcelColumnWidth LongWidth { get; set; } = ExcelColumnWidth.Characters(56);      // Evidence/Record/URL/Summary

        /// <summary>
        /// Case-insensitive set of header names that should use <see cref="ShortWidth"/>.
        /// </summary>
        public HashSet<string> ShortHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Case-insensitive set of header names that should use <see cref="NumericWidth"/>.
        /// </summary>
        public HashSet<string> NumericHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Case-insensitive set of header names that should use <see cref="MediumWidth"/>.
        /// </summary>
        public HashSet<string> MediumHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Case-insensitive set of header names that should use <see cref="LongWidth"/>.
        /// </summary>
        public HashSet<string> LongHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Case-insensitive set of header names that should be wrapped regardless of width.
        /// </summary>
        public HashSet<string> WrapHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Default character or physical width used for headers registered in <see cref="WrapHeaders"/>.
        /// </summary>
        public ExcelColumnWidth WrapWidth { get; set; } = ExcelColumnWidth.Characters(32);

        /// <summary>
        /// Headers that should be auto-fit instead of using the preset widths.
        /// </summary>
        public HashSet<string> AutoFitHeaders { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// When true, headers that do not have an explicit width/preset will be auto-fit.
        /// </summary>
        public bool AutoFitRemainingColumns { get; set; } = false;

        /// <summary>
        /// Explicit column width overrides keyed by header name. Values are character or physical dimensions.
        /// </summary>
        public Dictionary<string, ExcelColumnWidth> WidthByHeader { get; } = new(StringComparer.OrdinalIgnoreCase);
    }
}
