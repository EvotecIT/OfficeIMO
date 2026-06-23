namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Records the source BIFF record for decoded worksheet-level legacy metadata.
    /// </summary>
    public sealed class LegacyXlsWorksheetMetadataRecord {
        /// <summary>
        /// Creates a worksheet metadata provenance record.
        /// </summary>
        /// <param name="kind">Decoded metadata kind.</param>
        /// <param name="recordOffset">Byte offset of the source BIFF record.</param>
        /// <param name="recordType">BIFF record type identifier.</param>
        public LegacyXlsWorksheetMetadataRecord(LegacyXlsWorksheetMetadataKind kind, int recordOffset, ushort recordType) {
            Kind = kind;
            RecordOffset = recordOffset;
            RecordType = recordType;
        }

        /// <summary>Gets the decoded metadata kind.</summary>
        public LegacyXlsWorksheetMetadataKind Kind { get; }

        /// <summary>Gets the byte offset of the source BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }
    }
}
