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
            int recordOffset)
            : this(
                index,
                fillPattern,
                fillForegroundColor,
                fillBackgroundColor,
                fontColor: null,
                fontBold: null,
                fontItalic: null,
                recordType,
                recordOffset) {
        }

        /// <summary>
        /// Creates a legacy differential format model.
        /// </summary>
        public LegacyXlsDifferentialFormat(
            int index,
            byte? fillPattern,
            string? fillForegroundColor,
            string? fillBackgroundColor,
            string? fontColor,
            bool? fontBold,
            bool? fontItalic,
            ushort recordType,
            int recordOffset)
            : this(
                index,
                fillPattern,
                fillForegroundColor,
                fillBackgroundColor,
                fontColor,
                fontBold,
                fontItalic,
                recordType,
                recordOffset,
                border: null) {
        }

        /// <summary>
        /// Creates a legacy differential format model.
        /// </summary>
        public LegacyXlsDifferentialFormat(
            int index,
            byte? fillPattern,
            string? fillForegroundColor,
            string? fillBackgroundColor,
            string? fontColor,
            bool? fontBold,
            bool? fontItalic,
            ushort recordType,
            int recordOffset,
            LegacyXlsDifferentialBorder? border)
            : this(
                index,
                fillPattern,
                fillForegroundColor,
                fillBackgroundColor,
                fontColor,
                fontBold,
                fontItalic,
                recordType,
                recordOffset,
                border,
                numberFormatId: null,
                numberFormatCode: null) {
        }

        /// <summary>
        /// Creates a legacy differential format model.
        /// </summary>
        public LegacyXlsDifferentialFormat(
            int index,
            byte? fillPattern,
            string? fillForegroundColor,
            string? fillBackgroundColor,
            string? fontColor,
            bool? fontBold,
            bool? fontItalic,
            ushort recordType,
            int recordOffset,
            LegacyXlsDifferentialBorder? border,
            ushort? numberFormatId,
            string? numberFormatCode) {
            Index = index;
            FillPattern = fillPattern;
            FillForegroundColor = fillForegroundColor;
            FillBackgroundColor = fillBackgroundColor;
            FontColor = fontColor;
            FontBold = fontBold;
            FontItalic = fontItalic;
            RecordType = recordType;
            RecordOffset = recordOffset;
            Border = border;
            NumberFormatId = numberFormatId;
            NumberFormatCode = numberFormatCode;
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
        /// Gets the text color as ARGB hex, when present.
        /// </summary>
        public string? FontColor { get; }

        /// <summary>
        /// Gets whether the differential format explicitly applies bold text.
        /// </summary>
        public bool? FontBold { get; }

        /// <summary>
        /// Gets whether the differential format explicitly applies italic text.
        /// </summary>
        public bool? FontItalic { get; }

        /// <summary>
        /// Gets decoded border formatting, when present.
        /// </summary>
        public LegacyXlsDifferentialBorder? Border { get; }

        /// <summary>
        /// Gets the decoded number format id, when present.
        /// </summary>
        public ushort? NumberFormatId { get; }

        /// <summary>
        /// Gets the decoded number format code, when present.
        /// </summary>
        public string? NumberFormatCode { get; }

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
