namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Describes preserve-only non-compound feature metadata discovered while importing a legacy DOC document.
    /// </summary>
    public sealed class LegacyDocPreservedFeature {
        /// <summary>
        /// Creates preserved feature metadata.
        /// </summary>
        /// <param name="kind">Structured preserved feature category.</param>
        /// <param name="code">Stable feature code.</param>
        /// <param name="description">Human-readable feature description.</param>
        /// <param name="detailCode">Stable feature subtype key for reports and future import planning.</param>
        public LegacyDocPreservedFeature(
            LegacyDocPreservedFeatureKind kind,
            string code,
            string description,
            string detailCode) {
            Kind = kind;
            Code = code ?? throw new ArgumentNullException(nameof(code));
            Description = description ?? throw new ArgumentNullException(nameof(description));
            DetailCode = detailCode ?? throw new ArgumentNullException(nameof(detailCode));
        }

        /// <summary>
        /// Gets the structured preserved feature category.
        /// </summary>
        public LegacyDocPreservedFeatureKind Kind { get; }

        /// <summary>
        /// Gets the stable feature code.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Gets the human-readable feature description.
        /// </summary>
        public string Description { get; }

        /// <summary>
        /// Gets a stable feature subtype key for reports and future import planning.
        /// </summary>
        public string DetailCode { get; }
    }
}
