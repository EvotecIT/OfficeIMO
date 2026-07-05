namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>
    /// Describes preserve-only compound storage discovered while importing a legacy DOC document.
    /// </summary>
    public sealed class LegacyDocCompoundFeature {
        /// <summary>
        /// Creates preserved compound feature metadata.
        /// </summary>
        /// <param name="kind">Structured compound feature category.</param>
        /// <param name="code">Stable feature code.</param>
        /// <param name="description">Human-readable feature description.</param>
        /// <param name="entryPath">First compound-file entry path associated with the feature.</param>
        /// <param name="detailCode">Stable feature subtype key for reports and future import planning.</param>
        /// <param name="entryCount">Number of matching compound-file entries.</param>
        /// <param name="totalBytes">Sum of matching entry sizes, when reported by the compound container.</param>
        public LegacyDocCompoundFeature(
            LegacyDocCompoundFeatureKind kind,
            string code,
            string description,
            string? entryPath,
            string detailCode,
            int entryCount,
            long totalBytes) {
            Kind = kind;
            Code = code ?? throw new ArgumentNullException(nameof(code));
            Description = description ?? throw new ArgumentNullException(nameof(description));
            EntryPath = entryPath;
            DetailCode = detailCode ?? throw new ArgumentNullException(nameof(detailCode));
            EntryCount = entryCount;
            TotalBytes = totalBytes;
        }

        /// <summary>
        /// Gets the structured compound feature category.
        /// </summary>
        public LegacyDocCompoundFeatureKind Kind { get; }

        /// <summary>
        /// Gets the stable feature code.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Gets the human-readable feature description.
        /// </summary>
        public string Description { get; }

        /// <summary>
        /// Gets the first compound-file entry path associated with the feature, when known.
        /// </summary>
        public string? EntryPath { get; }

        /// <summary>
        /// Gets a stable feature subtype key for reports and future import planning.
        /// </summary>
        public string DetailCode { get; }

        /// <summary>
        /// Gets the number of matching compound-file entries.
        /// </summary>
        public int EntryCount { get; }

        /// <summary>
        /// Gets the sum of matching entry sizes, when reported by the compound container.
        /// </summary>
        public long TotalBytes { get; }
    }
}
