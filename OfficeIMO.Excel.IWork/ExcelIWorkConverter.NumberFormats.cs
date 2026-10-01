using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    private static string NumberFormatCode(IWorkNumberFormat format) {
        string number = format.ThousandsSeparator ? "#,##0" : "0";
        number += format.DecimalPlaces is int places
            ? places == 0 ? "" : "." + new string('0', places)
            : "." + new string('#', 15);
        if (format.Kind == IWorkNumberFormatKind.Percentage) number += "%";
        if (format.Kind == IWorkNumberFormatKind.Currency) {
            // The source identifier is qualified; a locale-specific currency symbol is not.
            // Keep the identifier visible and report the display approximation separately.
            string prefix = "\"" + format.CurrencyCode + " \"";
            if (format.UseAccountingStyle) return prefix + number + ";" + prefix + "(" + number + ")";
            string negative = format.NegativeStyle switch {
                IWorkNegativeNumberStyle.Red => "[Red]" + prefix + number,
                IWorkNegativeNumberStyle.Parentheses => prefix + "(" + number + ")",
                IWorkNegativeNumberStyle.RedAndParentheses => "[Red]" + prefix + "(" + number + ")",
                _ => prefix + "-" + number
            };
            return prefix + number + ";" + negative;
        }
        return format.NegativeStyle switch {
            IWorkNegativeNumberStyle.Red => number + ";[Red]" + number,
            IWorkNegativeNumberStyle.Parentheses => number + ";(" + number + ")",
            IWorkNegativeNumberStyle.RedAndParentheses => number + ";[Red](" + number + ")",
            _ => number
        };
    }

    private static IEnumerable<IWorkDiagnostic> CurrencyFormatDiagnostics(IWorkNumbersProjection projection) {
        if (projection.Sheets.SelectMany(sheet => sheet.Tables).SelectMany(table => table.Cells)
            .Any(cell => cell.NumberFormat?.Kind == IWorkNumberFormatKind.Currency)) {
            yield return new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_NUMBERS_CURRENCY_DISPLAY_APPROXIMATED",
                "Currency amounts retain their source three-letter identifier as a visible XLSX prefix. Currency symbols, locale-specific placement and accounting alignment are not reconstructed. Supported decimals, grouping and negative-value treatment are retained; values and formula caches remain numeric.",
                lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation);
        }
    }
}
