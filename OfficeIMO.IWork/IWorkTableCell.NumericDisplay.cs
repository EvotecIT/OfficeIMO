using OfficeIMO.Spreadsheet;

namespace OfficeIMO.IWork;

public sealed partial class IWorkTableCell {
    // Text-only document tables use the same numeric engine as spreadsheet views.
    // Keep raw values, caches and raw DisplayText properties unchanged on the source model.
    internal bool TryGetFormattedNumber(out string text, out string? color) {
        text = Kind == IWorkCellKind.Formula && Value != null ? CachedDisplayText : DisplayText;
        color = null;
        if (NumberFormat == null || !CachedValueIsComplete || HasDecodeError || RichText is { Paragraphs.Count: > 0 }) return false;
        string code = NumberFormat.ToSpreadsheetFormatCode();
        if (NumberFormat.DateTimeFormat != null) {
            if (ValueKind != IWorkCellKind.DateTime || Value is not DateTime date) return false;
            string? display = SpreadsheetNumberFormatDisplay.FormatDateTimeValue(date, code);
            if (display == null) return false;
            text = display;
            return true;
        }
        if (Value is not double number) return false;
        if (NumberFormat.Kind == IWorkNumberFormatKind.Duration) {
            if (ValueKind != IWorkCellKind.Duration
                || !SpreadsheetNumberFormatDisplay.TryFormatElapsedValue(number / 86400d, code, out string duration)) return false;
            text = duration;
            return true;
        }
        if (ValueKind != IWorkCellKind.Number) return false;
        string? formatted = SpreadsheetNumberFormatDisplay.FormatNumericValue(number, code);
        if (formatted == null) return false;
        text = formatted;
        color = SpreadsheetNumberFormatDisplay.GetNumericFormatColor(number, code);
        return true;
    }
}
