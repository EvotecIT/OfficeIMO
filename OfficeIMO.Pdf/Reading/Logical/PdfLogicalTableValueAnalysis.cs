using System.Globalization;
using System.Text;

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
    /// <summary>Values are true/false or yes/no booleans.</summary>
    Boolean,
    /// <summary>Values are unambiguous dates or date-times with explicit years.</summary>
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
    private static readonly string[] DateHeaderHints = {
        "date", "time", "due", "created", "updated", "modified", "issued", "expiry", "expires", "start", "end"
    };

    private static readonly string[] UnambiguousDateTimeFormats = {
        "yyyy-MM-dd", "yyyy/MM/dd", "yyyy.MM.dd",
        "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss",
        "yyyy/MM/dd HH:mm", "yyyy/MM/dd HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm", "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK"
    };

    /// <summary>Infers typed value profiles for normalized table data.</summary>
    public static IReadOnlyList<PdfLogicalTableValueProfile> Analyze(PdfLogicalTableData data, CultureInfo? culture = null) {
        Guard.NotNull(data, nameof(data));
        return Analyze(data.Columns, data.Rows, culture);
    }

    /// <summary>Infers typed value profiles for normalized columns and body rows.</summary>
    public static IReadOnlyList<PdfLogicalTableValueProfile> Analyze(
        IReadOnlyList<string> columns,
        IReadOnlyList<IReadOnlyList<string>> rows,
        CultureInfo? culture = null) {
        Guard.NotNull(columns, nameof(columns));
        Guard.NotNull(rows, nameof(rows));
        CultureInfo parsingCulture = culture ?? CultureInfo.InvariantCulture;
        var profiles = new PdfLogicalTableValueProfile[columns.Count];
        for (int columnIndex = 0; columnIndex < columns.Count; columnIndex++) {
            List<string> values = rows
                .Select(row => columnIndex < row.Count ? row[columnIndex].Trim() : string.Empty)
                .Where(static value => value.Length > 0)
                .ToList();
            PdfLogicalTableValueKind kind = InferKind(columns[columnIndex], values, parsingCulture);
            int matches = values.Count(value => Matches(kind, value, parsingCulture));
            profiles[columnIndex] = new PdfLogicalTableValueProfile(columnIndex, columns[columnIndex], kind, values.Count, matches);
        }
        return Array.AsReadOnly(profiles);
    }

    private static PdfLogicalTableValueKind InferKind(string columnName, List<string> values, CultureInfo culture) {
        if (values.Count == 0) return PdfLogicalTableValueKind.Empty;
        if (values.All(static value => TryParseBoolean(value, out _))) return PdfLogicalTableValueKind.Boolean;
        if (values.All(value => TryParsePercentage(value, culture, out _))) return PdfLogicalTableValueKind.Percentage;
        if (values.All(value => TryParseTime(value, culture, out _))) return PdfLogicalTableValueKind.Time;
        if (!values.Any(static value => LooksLikeAmbiguousNumericDate(value)) &&
            values.All(PdfLogicalTableAnalysis.LooksLikeNumericValue) &&
            values.All(value => PdfLogicalTableAnalysis.TryParseNumericValue(value, culture, out _))) return PdfLogicalTableValueKind.Number;
        if (HasDateSignal(columnName, values) &&
            values.All(static value => HasExplicitYear(value)) &&
            values.All(value => DateTime.TryParse(value, culture, DateTimeStyles.AllowWhiteSpaces, out _))) return PdfLogicalTableValueKind.DateTime;
        return PdfLogicalTableValueKind.Text;
    }

    private static bool Matches(PdfLogicalTableValueKind kind, string value, CultureInfo culture) => kind switch {
        PdfLogicalTableValueKind.Empty => false,
        PdfLogicalTableValueKind.Text => !LooksLikeTypedValue(value, culture),
        PdfLogicalTableValueKind.Boolean => TryParseBoolean(value, out _),
        PdfLogicalTableValueKind.Percentage => TryParsePercentage(value, culture, out _),
        PdfLogicalTableValueKind.Time => TryParseTime(value, culture, out _),
        PdfLogicalTableValueKind.Number => PdfLogicalTableAnalysis.TryParseNumericValue(value, culture, out _),
        PdfLogicalTableValueKind.DateTime => DateTime.TryParse(value, culture, DateTimeStyles.AllowWhiteSpaces, out _),
        _ => false
    };

    private static bool LooksLikeTypedValue(string value, CultureInfo culture) =>
        TryParseBoolean(value, out _) ||
        TryParsePercentage(value, culture, out _) ||
        TryParseTime(value, culture, out _) ||
        PdfLogicalTableAnalysis.TryParseNumericValue(value, culture, out _) ||
        (HasExplicitYear(value) && DateTime.TryParse(value, culture, DateTimeStyles.AllowWhiteSpaces, out _));

    private static bool TryParseBoolean(string value, out bool result) {
        string normalized = value.Trim();
        if (string.Equals(normalized, "true", StringComparison.OrdinalIgnoreCase) || string.Equals(normalized, "yes", StringComparison.OrdinalIgnoreCase)) { result = true; return true; }
        if (string.Equals(normalized, "false", StringComparison.OrdinalIgnoreCase) || string.Equals(normalized, "no", StringComparison.OrdinalIgnoreCase)) { result = false; return true; }
        result = false;
        return false;
    }

    private static bool TryParsePercentage(string value, CultureInfo culture, out decimal result) {
        string normalized = value.Trim();
        if (normalized.Length == 0 || normalized[normalized.Length - 1] != '%') { result = 0M; return false; }
        if (PdfLogicalTableAnalysis.TryParseNumericValue(normalized.Substring(0, normalized.Length - 1), culture, out decimal number)) { result = number / 100M; return true; }
        result = 0M;
        return false;
    }

    private static bool TryParseTime(string value, CultureInfo culture, out TimeSpan result) {
        string normalized = value.Trim();
        if (normalized.Length == 0 || normalized.IndexOf(':') < 0) { result = default; return false; }
        foreach (char current in normalized) {
            if (char.IsDigit(current) || char.IsWhiteSpace(current) || current is ':' or '.') continue;
            char upper = char.ToUpperInvariant(current);
            if (upper is 'A' or 'P' or 'M') continue;
            result = default;
            return false;
        }
        if (DateTime.TryParse(normalized, culture, DateTimeStyles.AllowWhiteSpaces, out DateTime parsed)) { result = parsed.TimeOfDay; return true; }
        result = default;
        return false;
    }

    private static bool HasDateSignal(string columnName, IReadOnlyList<string> values) {
        if (TokenizeHeaderWords(columnName).Any(word => DateHeaderHints.Contains(word, StringComparer.Ordinal))) return true;
        return values.All(static value => DateTime.TryParseExact(value.Trim(), UnambiguousDateTimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out _));
    }

    private static bool LooksLikeAmbiguousNumericDate(string value) {
        string normalized = value.Trim();
        char separator = normalized.Contains('/') ? '/' : normalized.Contains('-') ? '-' : normalized.Contains('.') ? '.' : '\0';
        if (separator == '\0') return false;
        string[] parts = normalized.Split(separator);
        return parts.Length == 3 && parts[2].Length == 4 &&
            int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out int first) &&
            int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out int second) &&
            int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out _) &&
            first is >= 1 and <= 12 && second is >= 1 and <= 12;
    }

    private static bool HasExplicitYear(string value) {
        int digitCount = 0;
        int number = 0;
        for (int index = 0; index <= value.Length; index++) {
            bool isDigit = index < value.Length && char.IsDigit(value[index]);
            if (isDigit) { digitCount++; number = digitCount <= 4 ? number * 10 + (value[index] - '0') : number; continue; }
            if (digitCount == 4 && number >= 1000) return true;
            digitCount = 0;
            number = 0;
        }
        return false;
    }

    private static IEnumerable<string> TokenizeHeaderWords(string value) {
        var word = new StringBuilder();
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (!char.IsLetterOrDigit(current)) {
                if (word.Length > 0) { yield return word.ToString().ToLowerInvariant(); word.Clear(); }
                continue;
            }
            if (word.Length > 0 && char.IsUpper(current) && char.IsLower(value[index - 1])) { yield return word.ToString().ToLowerInvariant(); word.Clear(); }
            word.Append(current);
        }
        if (word.Length > 0) yield return word.ToString().ToLowerInvariant();
    }
}
