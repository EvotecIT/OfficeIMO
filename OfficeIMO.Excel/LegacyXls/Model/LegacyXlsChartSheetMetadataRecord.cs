namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a decoded metadata record from a legacy chart-sheet substream.
    /// </summary>
    public sealed class LegacyXlsChartSheetMetadataRecord {
        internal LegacyXlsChartSheetMetadataRecord(
            LegacyXlsChartSheetMetadataKind kind,
            int recordOffset,
            ushort recordType) {
            Kind = kind;
            RecordOffset = recordOffset;
            RecordType = recordType;
        }

        /// <summary>
        /// Gets the metadata kind decoded from the source BIFF record.
        /// </summary>
        public LegacyXlsChartSheetMetadataKind Kind { get; }

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
