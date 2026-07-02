namespace OfficeIMO.Word.LegacyDoc {
    /// <summary>
    /// Controls legacy binary Word import behavior.
    /// </summary>
    public sealed class LegacyDocImportOptions {
        /// <summary>
        /// Maximum size, in bytes, of the extracted WordDocument stream.
        /// </summary>
        public int MaxWordDocumentStreamBytes { get; set; } = 64 * 1024 * 1024;

        /// <summary>
        /// When true, known unsupported legacy features are reported as diagnostics.
        /// </summary>
        public bool ReportUnsupportedFeatures { get; set; } = true;
    }
}
