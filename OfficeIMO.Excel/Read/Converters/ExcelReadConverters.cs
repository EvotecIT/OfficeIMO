using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Describes a cell's raw OpenXML data for custom conversion hooks.
    /// </summary>
    public readonly struct ExcelCellContext
    {
        public CellValues? TypeHint { get; }
        public uint? StyleIndex { get; }
        public string? RawText { get; }
        public string? InlineText { get; }
        public CultureInfo Culture { get; }

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
        public bool Handled { get; }
        public object? Value { get; }

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

        public static ExcelCellValue NotHandled => default;
    }
}

