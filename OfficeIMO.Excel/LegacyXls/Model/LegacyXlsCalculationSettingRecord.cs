namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Captures one calculation-related BIFF record and where it was found.
    /// </summary>
    public sealed class LegacyXlsCalculationSettingRecord {
        /// <summary>
        /// Creates a calculation setting record.
        /// </summary>
        public LegacyXlsCalculationSettingRecord(
            LegacyXlsCalculationSettingKind kind,
            string? sheetName,
            int recordOffset,
            ushort recordType,
            short? signedValue = null,
            double? doubleValue = null,
            bool? booleanValue = null,
            LegacyXlsCalculationMode? mode = null) {
            Kind = kind;
            SheetName = sheetName;
            RecordOffset = recordOffset;
            RecordType = recordType;
            SignedValue = signedValue;
            DoubleValue = doubleValue;
            BooleanValue = booleanValue;
            Mode = mode;
        }

        /// <summary>Gets the calculation setting kind.</summary>
        public LegacyXlsCalculationSettingKind Kind { get; }

        /// <summary>Gets the worksheet name when the record came from a sheet substream.</summary>
        public string? SheetName { get; }

        /// <summary>Gets the byte offset of the BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the signed integer value when the setting uses a 2-byte integer.</summary>
        public short? SignedValue { get; }

        /// <summary>Gets the floating-point value when the setting uses an 8-byte IEEE value.</summary>
        public double? DoubleValue { get; }

        /// <summary>Gets the Boolean value when the setting uses a BIFF Boolean flag.</summary>
        public bool? BooleanValue { get; }

        /// <summary>Gets the typed calculation mode when available.</summary>
        public LegacyXlsCalculationMode? Mode { get; }
    }
}
