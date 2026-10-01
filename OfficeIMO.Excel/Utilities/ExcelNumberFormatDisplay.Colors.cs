using OfficeIMO.Spreadsheet;

namespace OfficeIMO.Excel {
    internal static partial class ExcelNumberFormatDisplay {
        internal static string? GetNumericFormatColor(double value, uint numberFormatId, string? formatCode) {
            string? resolved = numberFormatId == 0U ? null : ResolveFormatCode(numberFormatId, formatCode);
            return string.IsNullOrWhiteSpace(resolved) ? null
                : SpreadsheetNumberFormatDisplay.GetNumericFormatColor(value, resolved!);
        }
    }
}
