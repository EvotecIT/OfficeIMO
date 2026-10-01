using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    private static string NumberFormatCode(IWorkNumberFormat format) {
        if (format.Kind == IWorkNumberFormatKind.Fraction) {
            string fraction = format.FractionAccuracy switch {
                IWorkFractionAccuracy.OneDigitDenominator => "# ?/?",
                IWorkFractionAccuracy.TwoDigitDenominator => "# ??/??",
                IWorkFractionAccuracy.ThreeDigitDenominator => "# ???/???",
                IWorkFractionAccuracy.Halves => "# ?/2",
                IWorkFractionAccuracy.Quarters => "# ?/4",
                IWorkFractionAccuracy.Eighths => "# ?/8",
                IWorkFractionAccuracy.Sixteenths => "# ??/16",
                IWorkFractionAccuracy.Tenths => "# ??/10",
                IWorkFractionAccuracy.Hundredths => "# ???/100",
                _ => throw new InvalidDataException("The fraction denominator precision is not supported.")
            };
            // Optional whole-number placeholders suppress zero in Excel. Keep
            // explicit minus and zero sections without adding a leading zero
            // to proper fractions.
            return fraction + ";-" + fraction + ";0";
        }
        string number = format.ThousandsSeparator ? "#,##0" : "0";
        number += format.DecimalPlaces is int places
            ? places == 0 ? "" : "." + new string('0', places)
            : "." + new string('#', 15);
        if (format.Kind == IWorkNumberFormatKind.Percentage) number += "%";
        if (format.Kind == IWorkNumberFormatKind.Scientific) number += "E+00";
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

    private static IEnumerable<IWorkDiagnostic> FractionFormatDiagnostics(IWorkNumbersProjection projection) {
        if (projection.Sheets.SelectMany(sheet => sheet.Tables).SelectMany(table => table.Cells)
            .Any(cell => cell.NumberFormat?.Kind == IWorkNumberFormatKind.Fraction)) {
            yield return new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_NUMBERS_FRACTION_DISPLAY_APPROXIMATED",
                "Fraction denominator precision is retained through Excel mixed-fraction formats. Midpoint rounding, equivalent-fraction normalization and spacing can differ from the source application; values and formula caches remain numeric.",
                lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation);
        }
    }
}
