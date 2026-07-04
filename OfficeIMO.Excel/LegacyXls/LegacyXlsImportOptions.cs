namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>
    /// Controls legacy binary Excel import behavior.
    /// </summary>
    public sealed class LegacyXlsImportOptions {
        /// <summary>
        /// Maximum size, in bytes, of the extracted BIFF workbook stream.
        /// </summary>
        public int MaxWorkbookStreamBytes { get; set; } = 64 * 1024 * 1024;

        /// <summary>
        /// When true, unsupported BIFF records are reported as warnings.
        /// </summary>
        public bool ReportUnsupportedRecords { get; set; } = true;

        /// <summary>
        /// Optional password used to decrypt password-to-open encrypted legacy XLS workbooks.
        /// </summary>
        public string? Password { get; set; }
    }
}
