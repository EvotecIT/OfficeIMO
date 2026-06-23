namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a hyperlink parsed from a legacy XLS worksheet.
    /// </summary>
    public sealed class LegacyXlsHyperlink {
        /// <summary>
        /// Creates a parsed legacy XLS hyperlink.
        /// </summary>
        public LegacyXlsHyperlink(int startRow, int startColumn, int endRow, int endColumn, string target, string? displayText = null, bool isExternal = true) {
            StartRow = startRow;
            StartColumn = startColumn;
            EndRow = endRow;
            EndColumn = endColumn;
            Target = target;
            DisplayText = displayText;
            IsExternal = isExternal;
        }

        /// <summary>Gets the first 1-based row covered by the hyperlink.</summary>
        public int StartRow { get; }

        /// <summary>Gets the first 1-based column covered by the hyperlink.</summary>
        public int StartColumn { get; }

        /// <summary>Gets the last 1-based row covered by the hyperlink.</summary>
        public int EndRow { get; }

        /// <summary>Gets the last 1-based column covered by the hyperlink.</summary>
        public int EndColumn { get; }

        /// <summary>Gets whether the hyperlink target is an external URI rather than an internal workbook location.</summary>
        public bool IsExternal { get; }

        /// <summary>Gets the hyperlink target URI or internal workbook location.</summary>
        public string Target { get; }

        /// <summary>Gets optional display text stored in the hyperlink object.</summary>
        public string? DisplayText { get; }
    }
}
