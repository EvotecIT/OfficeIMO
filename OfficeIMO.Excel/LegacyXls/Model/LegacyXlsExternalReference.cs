namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a legacy XLS supporting link discovered from a SupBook record.
    /// </summary>
    public sealed class LegacyXlsExternalReference {
        private readonly List<string> _sheetNames;
        private readonly List<LegacyXlsExternalName> _externalNames;
        private readonly List<LegacyXlsExternalCellCache> _cachedCellCaches;

        /// <summary>
        /// Creates supporting-link metadata.
        /// </summary>
        /// <param name="kind">Supporting-link category.</param>
        /// <param name="target">External target path or source name, when present.</param>
        /// <param name="sheetNames">External workbook sheet names, when present.</param>
        /// <param name="sheetCount">Sheet count declared by the SupBook record.</param>
        /// <param name="externalNames">External names declared after the supporting link, when present.</param>
        public LegacyXlsExternalReference(
            LegacyXlsExternalReferenceKind kind,
            string? target,
            IEnumerable<string>? sheetNames,
            ushort sheetCount,
            IEnumerable<LegacyXlsExternalName>? externalNames = null) {
            Kind = kind;
            Target = target;
            _sheetNames = sheetNames == null ? new List<string>() : new List<string>(sheetNames);
            _externalNames = externalNames == null ? new List<LegacyXlsExternalName>() : new List<LegacyXlsExternalName>(externalNames);
            _cachedCellCaches = new List<LegacyXlsExternalCellCache>();
            SheetCount = sheetCount;
        }

        /// <summary>
        /// Gets the supporting-link category.
        /// </summary>
        public LegacyXlsExternalReferenceKind Kind { get; }

        /// <summary>
        /// Gets the external target path or source name, when present.
        /// </summary>
        public string? Target { get; }

        /// <summary>
        /// Gets external workbook sheet names, when present.
        /// </summary>
        public IReadOnlyList<string> SheetNames => _sheetNames;

        /// <summary>
        /// Gets names declared by ExternName records following this supporting link.
        /// </summary>
        public IReadOnlyList<LegacyXlsExternalName> ExternalNames => _externalNames;

        /// <summary>
        /// Gets cached external cell values preserved from XCT/CRN record groups.
        /// </summary>
        public IReadOnlyList<LegacyXlsExternalCellCache> CachedCellCaches => _cachedCellCaches;

        /// <summary>
        /// Gets the sheet count declared by the SupBook record.
        /// </summary>
        public ushort SheetCount { get; }

        internal List<LegacyXlsExternalName> MutableExternalNames => _externalNames;

        internal List<LegacyXlsExternalCellCache> MutableCachedCellCaches => _cachedCellCaches;
    }
}
