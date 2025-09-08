using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Describes a cell's raw OpenXML data for custom conversion hooks.
    /// </summary>
    public readonly struct ExcelCellContext
    {
        /// <summary>OpenXML cell type hint, if present.</summary>
        public CellValues? TypeHint { get; }
        /// <summary>Cell style index as defined in the workbook, if present.</summary>
        public uint? StyleIndex { get; }
        /// <summary>Raw shared/inline text value when present.</summary>
        public string? RawText { get; }
        /// <summary>Inline string content when present.</summary>
        public string? InlineText { get; }
        /// <summary>Culture to use for parsing numbers/dates.</summary>
        public CultureInfo Culture { get; }

        /// <summary>
        /// Creates a description of the original OpenXML cell and culture for conversion.
        /// </summary>
        public ExcelCellContext(CellValues? typeHint, uint? styleIndex, string? rawText, string? inlineText, CultureInfo culture)
        {
            TypeHint = typeHint;
            StyleIndex = styleIndex;
            RawText = rawText;
            InlineText = inlineText;
            Culture = culture;
        }
    }

    /// <summary>
    /// Represents a custom-converted cell value. When Handled is true, Value should be used.
    /// When Handled is false, the built-in conversion pipeline should be used.
    /// </summary>
    public readonly struct ExcelCellValue
    {
        /// <summary>True when a custom converter produced a value and default conversion should be skipped.</summary>
        public bool Handled { get; }
        /// <summary>The converted value to use when <see cref="Handled"/> is true.</summary>
        public object? Value { get; }

        /// <summary>
        /// Creates a handled result carrying a custom <paramref name="value"/>.
        /// </summary>
        public ExcelCellValue(object? value)
        {
            Handled = true;
            Value = value;
        }

        private ExcelCellValue(bool handled, object? value)
        {
            Handled = handled;
            Value = value;
        }

        /// <summary>Signals that the custom converter did not handle the cell; continue with default conversion.</summary>
        public static ExcelCellValue NotHandled => default;
    }
}

