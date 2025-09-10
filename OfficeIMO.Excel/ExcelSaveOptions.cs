namespace OfficeIMO.Excel {
    /// <summary>
    /// Optional behaviors applied during <see cref="ExcelDocument.Save(string, bool, ExcelSaveOptions?)"/> and
    /// <see cref="ExcelDocument.SaveAsync(string, bool, ExcelSaveOptions?, System.Threading.CancellationToken)"/> to strengthen
    /// robustness and CI validation.
    /// </summary>
    public sealed class ExcelSaveOptions {
        /// <summary>
        /// When true, attempts to repair common defined-name issues (duplicates, out-of-range LocalSheetId, #REF!) before save.
        /// </summary>
        public bool SafeRepairDefinedNames { get; set; }

        /// <summary>
        /// When true, validates the saved package using <c>OpenXmlValidator</c> and throws on any errors.
        /// </summary>
        public bool ValidateOpenXml { get; set; }

        /// <summary>
        /// When true, performs a safety preflight on all worksheets before saving, removing empty containers
        /// (e.g., empty Hyperlinks/MergeCells), dropping orphaned drawing/header-footer references, and cleaning
        /// up invalid table references. This can prevent rare "Repaired Records" notices in Excel.
        /// </summary>
        public bool SafePreflight { get; set; }

        /// <summary>
        /// Returns an options instance with all features disabled.
        /// </summary>
        public static ExcelSaveOptions None => new ExcelSaveOptions();
    }
}
