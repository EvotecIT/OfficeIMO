using System.Globalization;

namespace OfficeIMO.Spreadsheet;

internal static partial class SpreadsheetNumberFormatDisplay {
    private static readonly DateTime Early1900Epoch = new DateTime(1899, 12, 31);

    /// <summary>
    /// Formats a 1900-date-system spreadsheet value, such as a cached chart value, as spreadsheet cells display it:
    /// numeric sections with conditions, fractions and exponents, and date, time and elapsed-time sections.
    /// Returns null when the format cannot represent the value.
    /// </summary>
    internal static string? FormatSerialValue(double value, string formatCode, IFormatProvider? provider = null) {
        if (double.IsNaN(value) || double.IsInfinity(value) || string.IsNullOrEmpty(formatCode)) return null;
        provider ??= CultureInfo.InvariantCulture;
        if (TryFormatElapsedValue(value, formatCode, out string elapsed)) return elapsed;
        string section = SelectNumberFormatSection(formatCode, value < 0 ? 1 : value == 0 ? 2 : 0, value, out _);
        if (HasElapsedTimeToken(section)) return null;
        if (!IsDateTimeSection(section)) return FormatNumericValue(value, formatCode, provider);
        return TryConvertSerial(value, out DateTime date) ? FormatDateTimeValue(date, formatCode, value, provider) : null;
    }

    /// <summary>Whether a format section shows a date or time rather than a number.</summary>
    internal static bool IsDateTimeSection(string section) {
        var syntax = SpreadsheetNumberFormatSyntax.Parse(section);
        return syntax.IsValid && syntax.Tokens.Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol) &&
            !syntax.Tokens.Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.Placeholder);
    }

    /// <summary>
    /// Converts a 1900-date-system serial. Serial 60 is Excel's fictitious February 29, 1900, shown as February 28;
    /// earlier serials count from December 31, 1899.
    /// </summary>
    internal static bool TryConvertSerial(double serial, out DateTime date) {
        date = default;
        if (double.IsNaN(serial) || serial < 0D || serial >= 2958466D) return false;
        date = serial < 60D ? Early1900Epoch.AddDays(serial) : DateTime.FromOADate(serial);
        return true;
    }
}
