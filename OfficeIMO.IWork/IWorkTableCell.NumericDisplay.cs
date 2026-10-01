using OfficeIMO.Spreadsheet;

namespace OfficeIMO.IWork;

public sealed partial class IWorkTableCell {
    // Text-only document tables use the same numeric engine as spreadsheet views.
    // Keep raw values, caches and raw DisplayText properties unchanged on the source model.
    internal bool TryGetFormattedNumber(out string text, out string? color) {
        text = Kind == IWorkCellKind.Formula && Value != null ? CachedDisplayText : DisplayText;
        color = null;
        if (NumberFormat == null || ValueKind != IWorkCellKind.Number || Value is not double number
            || !CachedValueIsComplete || HasDecodeError || RichText is { Paragraphs.Count: > 0 }) return false;
        string code = NumberFormat.ToSpreadsheetFormatCode();
        string? formatted = SpreadsheetNumberFormatDisplay.FormatNumericValue(number, code);
        if (formatted == null) return false;
        text = formatted;
        color = SpreadsheetNumberFormatDisplay.GetNumericFormatColor(number, code);
        return true;
    }
}
