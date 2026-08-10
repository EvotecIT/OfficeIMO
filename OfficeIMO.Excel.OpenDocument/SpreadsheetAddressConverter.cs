using System.Text;
using OfficeIMO.Spreadsheet;

namespace OfficeIMO.Excel.OpenDocument;

internal static class SpreadsheetAddressConverter {
    internal static SpreadsheetFormulaTranslationResult ExcelFormulaToOpenFormula(string formula) =>
        SpreadsheetFormulaSyntaxTree.Parse(formula ?? string.Empty, SpreadsheetFormulaDialect.ExcelA1)
            .TranslateTo(SpreadsheetFormulaDialect.OpenFormula);

    internal static SpreadsheetFormulaTranslationResult OpenFormulaToExcel(string formula) =>
        SpreadsheetFormulaSyntaxTree.Parse(formula ?? string.Empty, SpreadsheetFormulaDialect.OpenFormula)
            .TranslateTo(SpreadsheetFormulaDialect.ExcelA1);

    internal static string ExcelRangeToOpenAddress(string reference, string? defaultSheetName = null) {
        if (!SpreadsheetRangeReference.TryParse(reference, SpreadsheetAddressDialect.ExcelA1,
                out SpreadsheetRangeReference? parsed)) return string.Empty;
        string converted = parsed!.Format(SpreadsheetAddressDialect.OpenDocument);
        if (parsed.Start.SheetName != null || string.IsNullOrWhiteSpace(defaultSheetName)) return converted;

        string local = converted.StartsWith(".", StringComparison.Ordinal) ? converted.Substring(1) : converted;
        string escaped = defaultSheetName!.Replace("'", "''");
        return "$'" + escaped + "'." + local;
    }

    internal static string OpenAddressToExcel(string address) =>
        SpreadsheetRangeReference.TryParse(address, SpreadsheetAddressDialect.OpenDocument,
            out SpreadsheetRangeReference? parsed)
            && parsed!.TryFormat(SpreadsheetAddressDialect.ExcelA1, out string formatted)
            ? formatted
            : string.Empty;

    internal static string ToA1(int row, int column) {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        int value = column;
        var letters = new StringBuilder();
        while (value > 0) {
            value--;
            letters.Insert(0, (char)('A' + value % 26));
            value /= 26;
        }
        return letters.ToString() + row.ToString(CultureInfo.InvariantCulture);
    }
}