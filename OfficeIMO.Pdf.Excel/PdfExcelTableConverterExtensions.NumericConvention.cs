using System.Threading;
using System.Text.RegularExpressions;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Excel;

public static partial class PdfExcelTableConverterExtensions {
    private static readonly Regex NumericTokenPattern = new Regex(
        @"\d(?:[\d.,  ' ]*\d)?",
        RegexOptions.CultureInvariant,
        TimeSpan.FromSeconds(1));

    /// <summary>
    /// Infers whether the document writes decimals with a comma or a point from every imported table cell.
    /// A caller-selected culture other than the invariant default is kept, but contradicting evidence is reported.
    /// </summary>
    private static CultureInfo ResolveNumericCulture(
        IReadOnlyList<PdfCore.PdfLogicalTableContinuationGroup> tables,
        CultureInfo requested,
        List<OfficeConversionFidelityDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        NumericConventionEvidence evidence = default;
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalTableContinuationGroup table = tables[tableIndex];
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                IReadOnlyList<string> row = table.Rows[rowIndex];
                for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                    evidence = evidence.Add(ClassifyNumericSeparators(row[columnIndex]));
                }
            }
        }

        bool requestedIsDefault = ReferenceEquals(requested, CultureInfo.InvariantCulture) ||
            string.IsNullOrEmpty(requested.Name);
        char? inferredDecimal = evidence.Comma > 0 && evidence.Point == 0 ? ','
            : evidence.Point > 0 && evidence.Comma == 0 ? '.'
            : null;
        if (evidence.Comma > 0 && evidence.Point > 0) {
            diagnostics.Add(new OfficeConversionFidelityDiagnostic(
                "pdf-excel-numeric-convention-mixed",
                "PDF table values use both comma and point decimal separators (" +
                evidence.Comma.ToString(CultureInfo.InvariantCulture) + " comma, " +
                evidence.Point.ToString(CultureInfo.InvariantCulture) + " point); values were parsed with the " +
                DescribeCulture(requested) + " culture and some numbers may be misread.",
                OfficeConversionLossKind.Approximation,
                "OfficeIMO.Pdf.Excel"));
            return requested;
        }

        if (!inferredDecimal.HasValue) {
            if (evidence.Ambiguous > 0) {
                diagnostics.Add(new OfficeConversionFidelityDiagnostic(
                    "pdf-excel-numeric-value-ambiguous",
                    evidence.Ambiguous.ToString(CultureInfo.InvariantCulture) +
                    " PDF table value(s) such as '" + evidence.FirstAmbiguous + "' can be read either as a grouped integer or a decimal; " +
                    "the document has no other separator evidence, so they were parsed with the " + DescribeCulture(requested) + " culture.",
                    OfficeConversionLossKind.Approximation,
                    "OfficeIMO.Pdf.Excel"));
            }
            return requested;
        }

        string requestedDecimal = requested.NumberFormat.NumberDecimalSeparator;
        if (requestedDecimal.Length == 1 && requestedDecimal[0] == inferredDecimal.Value) return requested;
        if (!requestedIsDefault) {
            diagnostics.Add(new OfficeConversionFidelityDiagnostic(
                "pdf-excel-numeric-convention-mismatch",
                "PDF table values use '" + inferredDecimal.Value + "' as the decimal separator, but they were parsed with the requested " +
                DescribeCulture(requested) + " culture.",
                OfficeConversionLossKind.Approximation,
                "OfficeIMO.Pdf.Excel"));
            return requested;
        }

        return inferredDecimal.Value == ',' ? CreateDecimalCommaCulture() : CultureInfo.InvariantCulture;
    }

    private static string DescribeCulture(CultureInfo culture) =>
        string.IsNullOrEmpty(culture.Name) ? "invariant" : culture.Name;

    private static CultureInfo CreateDecimalCommaCulture() {
        var culture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        culture.NumberFormat.NumberDecimalSeparator = ",";
        culture.NumberFormat.NumberGroupSeparator = ".";
        culture.NumberFormat.CurrencyDecimalSeparator = ",";
        culture.NumberFormat.CurrencyGroupSeparator = ".";
        culture.NumberFormat.PercentDecimalSeparator = ",";
        culture.NumberFormat.PercentGroupSeparator = ".";
        return CultureInfo.ReadOnly(culture);
    }

    /// <summary>Classifies the separators of the single numeric token in a cell.</summary>
    internal static NumericSeparatorEvidence ClassifyNumericSeparators(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return NumericSeparatorEvidence.None;
        MatchCollection matches = NumericTokenPattern.Matches(value!);
        if (matches.Count != 1) return NumericSeparatorEvidence.None;
        string token = matches[0].Value;
        int lastComma = token.LastIndexOf(',');
        int lastPoint = token.LastIndexOf('.');
        if (lastComma < 0 && lastPoint < 0) return NumericSeparatorEvidence.None;
        if (lastComma >= 0 && lastPoint >= 0) {
            // Both separators: the last one is the decimal separator when the earlier ones are all the other kind.
            char decimalSeparator = lastComma > lastPoint ? ',' : '.';
            char groupSeparator = decimalSeparator == ',' ? '.' : ',';
            string integerPart = token.Substring(0, Math.Max(lastComma, lastPoint));
            if (integerPart.IndexOf(decimalSeparator) >= 0 || !HasThreeDigitGroups(integerPart, groupSeparator)) {
                return NumericSeparatorEvidence.None;
            }
            return decimalSeparator == ',' ? NumericSeparatorEvidence.CommaDecimal : NumericSeparatorEvidence.PointDecimal;
        }

        char separator = lastComma >= 0 ? ',' : '.';
        int count = 0;
        for (int index = 0; index < token.Length; index++) {
            if (token[index] == separator) count++;
        }
        if (count > 1) {
            // Repeated separators can only be grouping, which implies the other decimal convention.
            if (!HasThreeDigitGroups(token, separator)) return NumericSeparatorEvidence.None;
            return separator == ',' ? NumericSeparatorEvidence.PointDecimal : NumericSeparatorEvidence.CommaDecimal;
        }

        int digitsAfter = 0;
        for (int index = token.IndexOf(separator) + 1; index < token.Length; index++) {
            if (char.IsDigit(token[index])) digitsAfter++;
            else break;
        }
        if (digitsAfter == 3) {
            // "1,234" or "1.234" is a grouped integer in one convention and a decimal in the other.
            return NumericSeparatorEvidence.AmbiguousValue(token);
        }
        return separator == ',' ? NumericSeparatorEvidence.CommaDecimal : NumericSeparatorEvidence.PointDecimal;
    }

    private static bool HasThreeDigitGroups(string value, char groupSeparator) {
        string[] parts = value.Split(groupSeparator);
        if (parts.Length < 2) return true;
        string first = DigitsOnly(parts[0]);
        if (first.Length == 0 || first.Length > 3) return false;
        for (int index = 1; index < parts.Length; index++) {
            if (DigitsOnly(parts[index]).Length != 3) return false;
        }
        return true;
    }

    private static string DigitsOnly(string value) {
        var builder = new System.Text.StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (char.IsDigit(value[index])) builder.Append(value[index]);
        }
        return builder.ToString();
    }

    internal readonly struct NumericSeparatorEvidence {
        private NumericSeparatorEvidence(int comma, int point, string? ambiguous) {
            Comma = comma;
            Point = point;
            Ambiguous = ambiguous;
        }

        internal static NumericSeparatorEvidence None => default;
        internal static NumericSeparatorEvidence CommaDecimal => new NumericSeparatorEvidence(1, 0, null);
        internal static NumericSeparatorEvidence PointDecimal => new NumericSeparatorEvidence(0, 1, null);
        internal static NumericSeparatorEvidence AmbiguousValue(string token) => new NumericSeparatorEvidence(0, 0, token);

        internal int Comma { get; }
        internal int Point { get; }
        internal string? Ambiguous { get; }
    }

    private readonly struct NumericConventionEvidence {
        private NumericConventionEvidence(int comma, int point, int ambiguous, string? firstAmbiguous) {
            Comma = comma;
            Point = point;
            Ambiguous = ambiguous;
            FirstAmbiguous = firstAmbiguous;
        }

        internal int Comma { get; }
        internal int Point { get; }
        internal int Ambiguous { get; }
        internal string? FirstAmbiguous { get; }

        internal NumericConventionEvidence Add(NumericSeparatorEvidence value) => new NumericConventionEvidence(
            Comma + value.Comma,
            Point + value.Point,
            Ambiguous + (value.Ambiguous is null ? 0 : 1),
            FirstAmbiguous ?? value.Ambiguous);
    }
}
