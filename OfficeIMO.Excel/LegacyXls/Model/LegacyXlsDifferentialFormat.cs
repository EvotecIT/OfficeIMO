namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents differential formatting decoded from legacy XLS conditional-formatting style records.
    /// </summary>
    public sealed class LegacyXlsDifferentialFormat {
        /// <summary>
        /// Creates a legacy differential format model.
        /// </summary>
        public LegacyXlsDifferentialFormat(
            int index,
            byte? fillPattern,
            string? fillForegroundColor,
            string? fillBackgroundColor,
            ushort recordType,
            int recordOffset) {
            Index = index;
            FillPattern = fillPattern;
            FillForegroundColor = fillForegroundColor;
            FillBackgroundColor = fillBackgroundColor;
            RecordType = recordType;
            RecordOffset = recordOffset;
        }

        /// <summary>
        /// Gets the zero-based differential format index in the workbook collection.
        /// </summary>
        public int Index { get; }

        /// <summary>
        /// Gets the fill pattern code, when present.
        /// </summary>
        public byte? FillPattern { get; }

        /// <summary>
        /// Gets the foreground fill color as ARGB hex, when present.
        /// </summary>
        public string? FillForegroundColor { get; }

        /// <summary>
        /// Gets the background fill color as ARGB hex, when present.
        /// </summary>
        public string? FillBackgroundColor { get; }

        /// <summary>
        /// Gets the BIFF record type that supplied this differential format.
        /// </summary>
        public ushort RecordType { get; }

        /// <summary>
        /// Gets the BIFF stream offset of the source record.
        /// </summary>
        public int RecordOffset { get; }
    }
}
