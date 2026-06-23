namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a legacy XLS supporting link discovered from a SupBook record.
    /// </summary>
    public sealed class LegacyXlsExternalReference {
        private readonly List<string> _sheetNames;

        /// <summary>
        /// Creates supporting-link metadata.
        /// </summary>
        /// <param name="kind">Supporting-link category.</param>
        /// <param name="target">External target path or source name, when present.</param>
        /// <param name="sheetNames">External workbook sheet names, when present.</param>
        /// <param name="sheetCount">Sheet count declared by the SupBook record.</param>
        public LegacyXlsExternalReference(
            LegacyXlsExternalReferenceKind kind,
            string? target,
            IEnumerable<string>? sheetNames,
            ushort sheetCount) {
            Kind = kind;
            Target = target;
            _sheetNames = sheetNames == null ? new List<string>() : new List<string>(sheetNames);
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
        /// Gets the sheet count declared by the SupBook record.
        /// </summary>
        public ushort SheetCount { get; }
    }
}
