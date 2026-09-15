using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Typed value families inferred from normalized logical PDF table cells.</summary>
public enum PdfLogicalTableValueKind {
    /// <summary>No non-empty body values were observed.</summary>
    Empty,
    /// <summary>Values remain text.</summary>
    Text,
    /// <summary>Values are ordinary numbers.</summary>
    Number,
    /// <summary>Values are percentages with an explicit percent marker.</summary>
    Percentage,
    /// <summary>Values are invariant true/false Boolean literals.</summary>
    Boolean,
    /// <summary>Values are unambiguous dates or date-times, or dates parsed under an explicitly supplied culture.</summary>
    DateTime,
    /// <summary>Values are clock times without a date component.</summary>
    Time,
    /// <summary>Values are numbers with one consistent currency symbol or ISO code.</summary>
    Currency
}

/// <summary>Position of a detected currency affix relative to its numeric value.</summary>
public enum PdfLogicalCurrencyAffixPosition {
    /// <summary>The affix precedes the number.</summary>
    Prefix,
    /// <summary>The affix follows the number.</summary>
    Suffix
}

/// <summary>Shared typed-value evidence for one normalized table column.</summary>
public sealed class PdfLogicalTableValueProfile {
    internal PdfLogicalTableValueProfile(
        int index,
        string name,
        PdfLogicalTableValueKind kind,
        int nonEmptyCellCount,
        int matchingCellCount,
        string? currencyToken,
        PdfLogicalCurrencyAffixPosition? currencyAffixPosition,
        bool? currencyAffixUsesSpacing) {
        Index = index;
        Name = name ?? string.Empty;
        Kind = kind;
        NonEmptyCellCount = nonEmptyCellCount;
        MatchingCellCount = matchingCellCount;
        CurrencyToken = currencyToken;
        CurrencyAffixPosition = currencyAffixPosition;
        CurrencyAffixUsesSpacing = currencyAffixUsesSpacing;
        Confidence = nonEmptyCellCount == 0 ? 0D : (double)matchingCellCount / nonEmptyCellCount;
    }

    /// <summary>Zero-based normalized column index.</summary>
    public int Index { get; }
    /// <summary>Normalized column name.</summary>
    public string Name { get; }
    /// <summary>Inferred typed-value family.</summary>
    public PdfLogicalTableValueKind Kind { get; }
    /// <summary>Number of non-empty body cells inspected.</summary>
    public int NonEmptyCellCount { get; }
    /// <summary>Number of inspected cells matching <see cref="Kind"/>.</summary>
    public int MatchingCellCount { get; }
    /// <summary>Currency symbol or ISO code shared by the column when <see cref="Kind"/> is <see cref="PdfLogicalTableValueKind.Currency"/>.</summary>
    public string? CurrencyToken { get; }
    /// <summary>Consistent position of the currency affix relative to the number, or null for non-currency columns.</summary>
    public PdfLogicalCurrencyAffixPosition? CurrencyAffixPosition { get; }
    /// <summary>Whether source values consistently separate the currency affix from the number with whitespace, or null for non-currency columns.</summary>
    public bool? CurrencyAffixUsesSpacing { get; }
    /// <summary>Matching-cell ratio from 0 to 1.</summary>
    public double Confidence { get; }
}

/// <summary>Culture-aware typed-value inference shared by reverse-conversion adapters.</summary>
public static class PdfLogicalTableValueAnalysis {
    /// <summary>Infers typed value profiles for normalized table data.</summary>
    public static IReadOnlyList<PdfLogicalTableValueProfile> Analyze(
        PdfLogicalTableData data,
        PdfLogicalTableValueAnalysisOptions? options = null) {
        Guard.NotNull(data, nameof(data));
        return Analyze(data.Columns, data.Rows, options);
    }

    /// <summary>Infers typed value profiles for normalized columns and body rows.</summary>
    public static IReadOnlyList<PdfLogicalTableValueProfile> Analyze(
        IReadOnlyList<string> columns,
        IReadOnlyList<IReadOnlyList<string>> rows,
        PdfLogicalTableValueAnalysisOptions? options = null) {
        Guard.NotNull(columns, nameof(columns));
        Guard.NotNull(rows, nameof(rows));
        CultureInfo numericCulture = options?.NumericCulture ?? CultureInfo.InvariantCulture;
        CultureInfo? dateTimeCulture = options?.DateTimeCulture;
        var profiles = new PdfLogicalTableValueProfile[columns.Count];
        for (int columnIndex = 0; columnIndex < columns.Count; columnIndex++) {
            List<string> values = rows
                .Select(row => columnIndex < row.Count ? row[columnIndex].Trim() : string.Empty)
                .Where(static value => value.Length > 0)
                .ToList();
            PdfLogicalTableValueKind kind = InferKind(
                values,
                numericCulture,
                dateTimeCulture,
                out string? currencyToken,
                out PdfLogicalCurrencyAffixPosition? currencyAffixPosition,
                out bool? currencyAffixUsesSpacing);
            int matches = values.Count(value => Matches(kind, value, numericCulture, dateTimeCulture));
            profiles[columnIndex] = new PdfLogicalTableValueProfile(
                columnIndex,
                columns[columnIndex],
                kind,
                values.Count,
                matches,
                currencyToken,
                currencyAffixPosition,
                currencyAffixUsesSpacing);
        }
        return Array.AsReadOnly(profiles);
    }

    private static PdfLogicalTableValueKind InferKind(
        List<string> values,
        CultureInfo numericCulture,
        CultureInfo? dateTimeCulture,
        out string? currencyToken,
        out PdfLogicalCurrencyAffixPosition? currencyAffixPosition,
        out bool? currencyAffixUsesSpacing) {
        currencyToken = null;
        currencyAffixPosition = null;
        currencyAffixUsesSpacing = null;
        if (values.Count == 0) return PdfLogicalTableValueKind.Empty;
        if (values.All(static value => PdfLogicalTableValueParser.TryParseBoolean(value, out _))) return PdfLogicalTableValueKind.Boolean;
        if (values.All(value => PdfLogicalTableValueParser.TryParsePercentage(value, numericCulture, out _))) return PdfLogicalTableValueKind.Percentage;
        if (values.All(value => PdfLogicalTableValueParser.TryParseTime(value, dateTimeCulture, out _))) return PdfLogicalTableValueKind.Time;
        if (values.All(value => PdfLogicalTableValueParser.TryParseDateTime(value, dateTimeCulture, out _))) return PdfLogicalTableValueKind.DateTime;
        var currencyTokens = new string[values.Count];
        var currencyPositions = new PdfLogicalCurrencyAffixPosition[values.Count];
        var currencySpacing = new bool[values.Count];
        for (int index = 0; index < values.Count; index++) {
            if (!PdfLogicalTableValueParser.TryParseCurrency(
                    values[index],
                    numericCulture,
                    out _,
                    out currencyTokens[index],
                    out currencyPositions[index],
                    out currencySpacing[index])) {
                currencyTokens[index] = string.Empty;
            }
        }
        if (currencyTokens.All(static token => token.Length > 0)) {
            if (currencyTokens.Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1 &&
                currencyPositions.Distinct().Count() == 1 &&
                currencySpacing.Distinct().Count() == 1) {
                currencyToken = currencyTokens[0];
                currencyAffixPosition = currencyPositions[0];
                currencyAffixUsesSpacing = currencySpacing[0];
                return PdfLogicalTableValueKind.Currency;
            }

            return PdfLogicalTableValueKind.Text;
        }
        if (!values.Any(static value => PdfLogicalTableValueParser.LooksLikePlausibleNumericDate(value)) &&
            values.All(PdfLogicalTableAnalysis.LooksLikeNumericValue) &&
            values.All(value => PdfLogicalTableAnalysis.TryParseNumericValue(value, numericCulture, out _))) return PdfLogicalTableValueKind.Number;
        return PdfLogicalTableValueKind.Text;
    }

    private static bool Matches(
        PdfLogicalTableValueKind kind,
        string value,
        CultureInfo numericCulture,
        CultureInfo? dateTimeCulture) => kind switch {
        PdfLogicalTableValueKind.Empty => false,
        PdfLogicalTableValueKind.Text => !LooksLikeTypedValue(value, numericCulture, dateTimeCulture),
        PdfLogicalTableValueKind.Boolean => PdfLogicalTableValueParser.TryParseBoolean(value, out _),
        PdfLogicalTableValueKind.Percentage => PdfLogicalTableValueParser.TryParsePercentage(value, numericCulture, out _),
        PdfLogicalTableValueKind.Time => PdfLogicalTableValueParser.TryParseTime(value, dateTimeCulture, out _),
        PdfLogicalTableValueKind.Number => PdfLogicalTableAnalysis.TryParseNumericValue(value, numericCulture, out _),
        PdfLogicalTableValueKind.Currency => PdfLogicalTableValueParser.TryParseCurrency(value, numericCulture, out _, out _),
        PdfLogicalTableValueKind.DateTime => PdfLogicalTableValueParser.TryParseDateTime(value, dateTimeCulture, out _),
        _ => false
    };

    private static bool LooksLikeTypedValue(
        string value,
        CultureInfo numericCulture,
        CultureInfo? dateTimeCulture) =>
        PdfLogicalTableValueParser.TryParseBoolean(value, out _) ||
        PdfLogicalTableValueParser.TryParsePercentage(value, numericCulture, out _) ||
        PdfLogicalTableValueParser.TryParseTime(value, dateTimeCulture, out _) ||
        PdfLogicalTableValueParser.TryParseDateTime(value, dateTimeCulture, out _) ||
        PdfLogicalTableValueParser.TryParseCurrency(value, numericCulture, out _, out _) ||
        (!PdfLogicalTableValueParser.LooksLikePlausibleNumericDate(value) &&
         PdfLogicalTableAnalysis.TryParseNumericValue(value, numericCulture, out _));

}
