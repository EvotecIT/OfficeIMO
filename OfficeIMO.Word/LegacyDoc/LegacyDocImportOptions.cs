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

        /// <summary>Maximum number of text characters decoded from the legacy piece table.</summary>
        public int MaxDecodedCharacters { get; set; } = 16 * 1024 * 1024;

        /// <summary>Maximum aggregate paragraph-property traversal and expansion work, in bytes, per import.</summary>
        /// <remarks>
        /// Data pointers consume work even when they emit no properties. Cached expansions also consume work.
        /// Exhaustion stops the affected property expansion and reports a DOC-PAPX-INVALID diagnostic.
        /// </remarks>
        public int MaxParagraphPropertyWorkBytes { get; set; } = 64 * 1024 * 1024;

        /// <summary>Maximum number of Data records followed by one paragraph-property chain.</summary>
        /// <remarks>Cached suffixes retain their original depth. Exceeding this limit reports a DOC-PAPX-INVALID diagnostic.</remarks>
        public int MaxDataPropertyChainLength { get; set; } = 256;

        /// <summary>
        /// When true, known unsupported legacy content is reported as diagnostics.
        /// </summary>
        public bool ReportUnsupportedContent { get; set; } = true;

        internal void Validate() {
            if (MaxInputBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
            if (MaxDecodedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDecodedImageBytes));
            if (MaxDecodedCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDecodedCharacters));
            if (MaxParagraphPropertyWorkBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxParagraphPropertyWorkBytes));
            if (MaxDataPropertyChainLength <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDataPropertyChainLength));
        }

    }
}
