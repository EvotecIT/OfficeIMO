namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents the workbook-level BIFF TableStyles record.
    /// </summary>
    public sealed class LegacyXlsTableStyleCollection {
        /// <summary>
        /// Initializes a new instance of the <see cref="LegacyXlsTableStyleCollection"/> class.
        /// </summary>
        public LegacyXlsTableStyleCollection(
            uint totalStyleCount,
            string? defaultTableStyleName,
            string? defaultPivotStyleName,
            ushort headerRecordType,
            ushort headerFlags,
            int recordOffset,
            ushort recordType,
            int payloadLength) {
            TotalStyleCount = totalStyleCount;
            DefaultTableStyleName = defaultTableStyleName;
            DefaultPivotStyleName = defaultPivotStyleName;
            HeaderRecordType = headerRecordType;
            HeaderFlags = headerFlags;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
        }

        /// <summary>Gets the total built-in plus custom table style count declared by the workbook.</summary>
        public uint TotalStyleCount { get; }

        /// <summary>Gets the workbook default table style name, when declared.</summary>
        public string? DefaultTableStyleName { get; }

        /// <summary>Gets the workbook default PivotTable style name, when declared.</summary>
        public string? DefaultPivotStyleName { get; }

        /// <summary>Gets the FRT header record type stored inside the payload.</summary>
        public ushort HeaderRecordType { get; }

        /// <summary>Gets the FRT header flags stored inside the payload.</summary>
        public ushort HeaderFlags { get; }

        /// <summary>Gets the BIFF stream offset of the source record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the source BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the source BIFF payload length.</summary>
        public int PayloadLength { get; }
    }
}
