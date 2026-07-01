namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a BIFF DConName record that points a consolidation source at a defined name.
    /// </summary>
    public sealed class LegacyXlsDataConsolidationName {
        /// <summary>
        /// Creates decoded DConName metadata.
        /// </summary>
        public LegacyXlsDataConsolidationName(
            int recordOffset,
            ushort recordType,
            string name,
            LegacyXlsDataConsolidationSourceKind sourceKind,
            string source,
            int unusedByteCount) {
            RecordOffset = recordOffset;
            RecordType = recordType;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            SourceKind = sourceKind;
            Source = source ?? throw new ArgumentNullException(nameof(source));
            UnusedByteCount = unusedByteCount;
        }

        /// <summary>Gets the byte offset of the DConName BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the workbook or external defined-name reference.</summary>
        public string Name { get; }

        /// <summary>Gets the decoded source workbook shape.</summary>
        public LegacyXlsDataConsolidationSourceKind SourceKind { get; }

        /// <summary>Gets the external source string, or an empty string for workbook-scoped names.</summary>
        public string Source { get; }

        /// <summary>Gets the count of unused trailing bytes after the DConName payload.</summary>
        public int UnusedByteCount { get; }
    }
}
