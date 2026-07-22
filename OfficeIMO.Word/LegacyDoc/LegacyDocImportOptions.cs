namespace OfficeIMO.Word.LegacyDoc {
    /// <summary>
    /// Controls legacy binary Word import behavior.
    /// </summary>
    public sealed class LegacyDocImportOptions {
        /// <summary>
        /// Maximum size, in bytes, of the extracted document input stream.
        /// </summary>
        public int MaxInputBytes { get; set; } = 64 * 1024 * 1024;

        /// <summary>Maximum aggregate decoded OfficeArt image bytes retained during import.</summary>
        public int MaxDecodedImageBytes { get; set; } = 64 * 1024 * 1024;

        /// <summary>
        /// When true, known unsupported legacy content is reported as diagnostics.
        /// </summary>
        public bool ReportUnsupportedContent { get; set; } = true;

        internal void Validate() {
            if (MaxInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
            if (MaxDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDecodedImageBytes));
        }

    }
}
