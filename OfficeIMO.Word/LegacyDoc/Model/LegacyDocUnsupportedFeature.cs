namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Describes an unsupported or preserve-only feature discovered in a legacy DOC document.
    /// </summary>
    public sealed class LegacyDocUnsupportedFeature {
        /// <summary>
        /// Creates unsupported feature metadata.
        /// </summary>
        /// <param name="kind">Structured unsupported feature category.</param>
        /// <param name="code">Stable feature/diagnostic code.</param>
        /// <param name="description">Human-readable feature description.</param>
        /// <param name="entryPath">Compound-file entry path associated with the feature, when known.</param>
        /// <param name="detailCode">Stable feature subtype key for reports and future import planning.</param>
        public LegacyDocUnsupportedFeature(
            LegacyDocUnsupportedFeatureKind kind,
            string code,
            string description,
            string? entryPath = null,
            string? detailCode = null) {
            Kind = kind;
            Code = code ?? throw new ArgumentNullException(nameof(code));
            Description = description ?? throw new ArgumentNullException(nameof(description));
            EntryPath = entryPath;
            DetailCode = detailCode;
        }

        /// <summary>
        /// Gets the structured unsupported feature category.
        /// </summary>
        public LegacyDocUnsupportedFeatureKind Kind { get; }

        /// <summary>
        /// Gets the stable feature/diagnostic code.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Gets the human-readable feature description.
        /// </summary>
        public string Description { get; }

        /// <summary>
        /// Gets the compound-file entry path associated with the feature, when known.
        /// </summary>
        public string? EntryPath { get; }

        /// <summary>
        /// Gets a stable feature subtype key for reports and future import planning.
        /// </summary>
        public string? DetailCode { get; }
    }
}
