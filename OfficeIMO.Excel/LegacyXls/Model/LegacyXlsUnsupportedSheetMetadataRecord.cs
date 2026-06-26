namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a decoded metadata record from an unsupported legacy sheet substream.
    /// </summary>
    public sealed class LegacyXlsUnsupportedSheetMetadataRecord {
        internal LegacyXlsUnsupportedSheetMetadataRecord(
            LegacyXlsUnsupportedSheetMetadataKind kind,
            int recordOffset,
            ushort recordType) {
            Kind = kind;
            RecordOffset = recordOffset;
            RecordType = recordType;
        }

        /// <summary>
        /// Gets the metadata kind decoded from the source BIFF record.
        /// </summary>
        public LegacyXlsUnsupportedSheetMetadataKind Kind { get; }

        /// <summary>
        /// Gets the byte offset of the source BIFF record.
        /// </summary>
        public int RecordOffset { get; }

        /// <summary>
        /// Gets the source BIFF record type.
        /// </summary>
        public ushort RecordType { get; }
    }
}
