using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    private static IEnumerable<IWorkDiagnostic> DateTimeFormatDiagnostics(IWorkNumbersProjection projection) {
        if (projection.Sheets.SelectMany(sheet => sheet.Tables).SelectMany(table => table.Cells)
            .Any(cell => cell.NumberFormat?.Kind == IWorkNumberFormatKind.DateTime)) {
            yield return new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_NUMBERS_DATE_DISPLAY_APPROXIMATED",
                "Qualified source date/time patterns map to Excel calendar formats without changing dates or formula caches. Localized punctuation, month names and day-period labels can differ; calendar, locale and time-zone equivalence are unqualified. Source patterns remain on the projection. Existing XLSX date-range and precision guards still apply.",
                lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation);
        }
    }

    private static IEnumerable<IWorkDiagnostic> DurationFormatDiagnostics(IWorkNumbersProjection projection) {
        if (projection.Sheets.SelectMany(sheet => sheet.Tables).SelectMany(table => table.Cells)
            .Any(cell => cell.NumberFormat?.Kind == IWorkNumberFormatKind.Duration)) {
            yield return new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_NUMBERS_DURATION_DISPLAY_APPROXIMATED",
                "Fixed abbreviated hours and minutes use an elapsed XLSX format, retaining total hours, signs, numeric values and formula caches. Source seconds are converted to day serials. Unit spacing and rounding can differ; other duration settings and locale equivalence are unqualified.",
                lossKind: global::OfficeIMO.OfficeConversionLossKind.Approximation);
        }
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
