namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes an unsupported or preserve-only feature discovered in a legacy XLS workbook.
    /// </summary>
    public sealed class LegacyXlsUnsupportedFeature {
        /// <summary>
        /// Creates unsupported feature metadata.
        /// </summary>
        /// <param name="kind">Structured unsupported feature category.</param>
        /// <param name="code">Stable feature/diagnostic code.</param>
        /// <param name="description">Human-readable feature description.</param>
        /// <param name="sheetName">Worksheet or sheet entry name associated with the feature, when known.</param>
        /// <param name="recordOffset">Byte offset of the related BIFF record, when known.</param>
        /// <param name="recordType">BIFF record type identifier, when known.</param>
        /// <param name="detailCode">Stable feature subtype key for reports and future import planning.</param>
        public LegacyXlsUnsupportedFeature(
            LegacyXlsUnsupportedFeatureKind kind,
            string code,
            string description,
            string? sheetName = null,
            int? recordOffset = null,
            ushort? recordType = null,
            string? detailCode = null) {
            Kind = kind;
            Code = code ?? throw new ArgumentNullException(nameof(code));
            Description = description ?? throw new ArgumentNullException(nameof(description));
            SheetName = sheetName;
            RecordOffset = recordOffset;
            RecordType = recordType;
            DetailCode = detailCode;
        }

        /// <summary>
        /// Gets the structured unsupported feature category.
        /// </summary>
        public LegacyXlsUnsupportedFeatureKind Kind { get; }

        /// <summary>
        /// Gets the stable feature/diagnostic code.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Gets the human-readable feature description.
        /// </summary>
        public string Description { get; }

        /// <summary>
        /// Gets the worksheet or sheet entry name associated with the feature, when known.
        /// </summary>
        public string? SheetName { get; }

        /// <summary>
        /// Gets the byte offset of the related BIFF record, when known.
        /// </summary>
        public int? RecordOffset { get; }

        /// <summary>
        /// Gets the BIFF record type identifier, when known.
        /// </summary>
        public ushort? RecordType { get; }

        /// <summary>
        /// Gets a stable feature subtype key for reports and future import planning.
        /// </summary>
        public string? DetailCode { get; }
    }
}
