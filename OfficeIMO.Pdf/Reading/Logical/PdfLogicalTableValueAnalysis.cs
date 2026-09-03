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
    Time
}

/// <summary>Shared typed-value evidence for one normalized table column.</summary>
public sealed class PdfLogicalTableValueProfile {
    internal PdfLogicalTableValueProfile(int index, string name, PdfLogicalTableValueKind kind, int nonEmptyCellCount, int matchingCellCount) {
        Index = index;
        Name = name ?? string.Empty;
        Kind = kind;
        NonEmptyCellCount = nonEmptyCellCount;
        MatchingCellCount = matchingCellCount;
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
            PdfLogicalTableValueKind kind = InferKind(values, numericCulture, dateTimeCulture);
            int matches = values.Count(value => Matches(kind, value, numericCulture, dateTimeCulture));
            profiles[columnIndex] = new PdfLogicalTableValueProfile(columnIndex, columns[columnIndex], kind, values.Count, matches);
        }
        return Array.AsReadOnly(profiles);
    }

    private static PdfLogicalTableValueKind InferKind(
        List<string> values,
        CultureInfo numericCulture,
        CultureInfo? dateTimeCulture) {
        if (values.Count == 0) return PdfLogicalTableValueKind.Empty;
        if (values.All(static value => PdfLogicalTableValueParser.TryParseBoolean(value, out _))) return PdfLogicalTableValueKind.Boolean;
        if (values.All(value => PdfLogicalTableValueParser.TryParsePercentage(value, numericCulture, out _))) return PdfLogicalTableValueKind.Percentage;
        if (values.All(value => PdfLogicalTableValueParser.TryParseTime(value, dateTimeCulture, out _))) return PdfLogicalTableValueKind.Time;
        if (values.All(value => PdfLogicalTableValueParser.TryParseDateTime(value, dateTimeCulture, out _))) return PdfLogicalTableValueKind.DateTime;
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
        (!PdfLogicalTableValueParser.LooksLikePlausibleNumericDate(value) &&
         PdfLogicalTableAnalysis.TryParseNumericValue(value, numericCulture, out _));

}
