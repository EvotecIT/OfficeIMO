namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>
    /// Controls legacy binary Excel import behavior.
    /// </summary>
    public sealed class LegacyXlsImportOptions {
        /// <summary>
        /// Maximum size, in bytes, of the extracted workbook input stream.
        /// </summary>
        public int MaxInputBytes { get; set; } = 64 * 1024 * 1024;

        /// <summary>Maximum aggregate decoded OfficeArt image bytes retained during import.</summary>
        public int MaxDecodedImageBytes { get; set; } = 64 * 1024 * 1024;

        /// <summary>
        /// When true, unsupported legacy content is reported as warnings.
        /// </summary>
        public bool ReportUnsupportedContent { get; set; } = true;

        /// <summary>
        /// Optional password used to decrypt password-to-open encrypted legacy XLS workbooks.
        /// </summary>
        public string? Password { get; set; }

        internal void Validate() {
            if (MaxInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
            if (MaxDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDecodedImageBytes));
        }

    }
}
