namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Records the source BIFF record for decoded workbook-level legacy metadata.
    /// </summary>
    public sealed class LegacyXlsWorkbookMetadataRecord {
        /// <summary>
        /// Creates a workbook metadata provenance record.
        /// </summary>
        /// <param name="kind">Decoded metadata kind.</param>
        /// <param name="recordOffset">Byte offset of the source BIFF record.</param>
        /// <param name="recordType">BIFF record type identifier.</param>
        public LegacyXlsWorkbookMetadataRecord(LegacyXlsWorkbookMetadataKind kind, int recordOffset, ushort recordType) {
            Kind = kind;
            RecordOffset = recordOffset;
            RecordType = recordType;
        }

        /// <summary>Gets the decoded metadata kind.</summary>
        public LegacyXlsWorkbookMetadataKind Kind { get; }

        /// <summary>Gets the byte offset of the source BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }
    }
}
