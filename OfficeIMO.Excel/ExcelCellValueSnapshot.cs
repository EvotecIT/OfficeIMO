using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Identifies the native value kind stored in a worksheet cell.
    /// </summary>
    public enum ExcelCellValueKind {
        /// <summary>The cell stores text.</summary>
        Text,
        /// <summary>The cell stores a numeric value.</summary>
        Number,
        /// <summary>The cell stores a boolean value.</summary>
        Boolean,
        /// <summary>The cell stores an error value.</summary>
        Error,
        /// <summary>The cell stores a formula.</summary>
        Formula,
        /// <summary>The cell stores a value kind not explicitly mapped by OfficeIMO.</summary>
        Other,
        /// <summary>The cell stores a numeric Excel serial with a date-like number format.</summary>
        DateTime
    }

    /// <summary>
    /// Describes a native worksheet cell value without requiring callers to inspect OpenXML directly.
    /// </summary>
    public sealed class ExcelCellValueSnapshot {
        internal ExcelCellValueSnapshot(ExcelCellValueKind kind, string text, string rawValue, CellValues? openXmlType, DateTime? dateTimeValue = null) {
            Kind = kind;
            Text = text;
            RawValue = rawValue;
            OpenXmlType = openXmlType;
            DateTimeValue = dateTimeValue;
        }

        /// <summary>Native value kind for the cell.</summary>
        public ExcelCellValueKind Kind { get; }

        /// <summary>Resolved display text for the cell.</summary>
        public string Text { get; }

        /// <summary>Raw value that can be used for loss-aware interchange between OfficeIMO converters.</summary>
        public string RawValue { get; }

        /// <summary>Underlying OpenXML cell type hint, when present.</summary>
        public CellValues? OpenXmlType { get; }

        /// <summary>
        /// Resolved date/time value when <see cref="Kind"/> is <see cref="ExcelCellValueKind.DateTime"/>.
        /// </summary>
        public DateTime? DateTimeValue { get; }
    }
}
