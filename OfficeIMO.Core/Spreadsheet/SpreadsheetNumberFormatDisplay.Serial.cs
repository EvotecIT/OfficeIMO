using System.Globalization;

namespace OfficeIMO.Spreadsheet;

internal static partial class SpreadsheetNumberFormatDisplay {
    private static readonly DateTime Early1900Epoch = new DateTime(1899, 12, 31);

    /// <summary>What a date or time section shows for a serial outside the 1900 date system, such as a negative value.</summary>
    internal const string InvalidDateDisplay = "#####";

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
        return TryConvertSerial(value, out DateTime date) ? FormatDateTimeValue(date, formatCode, value, provider) : InvalidDateDisplay;
    }

    /// <summary>Whether a format section shows a date or time rather than a number.</summary>
    internal static bool IsDateTimeSection(string section) {
        var syntax = SpreadsheetNumberFormatSyntax.Parse(section);
        if (!syntax.IsValid || !syntax.Tokens.Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol)) return false;
        // Digits after the seconds, as in hh:mm:ss.00, are fractional seconds rather than a number.
        for (int index = 0; index < syntax.Tokens.Count; index++)
            if (syntax.Tokens[index].Kind == SpreadsheetNumberFormatTokenKind.Placeholder && !IsFractionalSecond(syntax.Tokens, index)) return false;
        return true;
    }

    /// <summary>
    /// Converts a 1900-date-system serial. Serial 60 is Excel's fictitious February 29, 1900, shown as February 28;
    /// earlier serials count from December 31, 1899.
    /// </summary>
    internal static bool TryConvertSerial(double serial, out DateTime date) {
        date = default;
        if (double.IsNaN(serial) || serial < 0D || serial >= 2958466D) return false;
        try {
            // FromOADate rounds to the millisecond, so a serial just below the limit can still overflow.
            date = serial < 60D ? Early1900Epoch.AddDays(serial) : DateTime.FromOADate(serial);
        } catch (ArgumentException) {
            return false;
        }
        return true;
    }

    /// <summary>
    /// Rounds a value to the displayed decimals as spreadsheets do, from its 15-significant-digit decimal form, so
    /// 1.005 with two decimals shows as 1.01 although the stored binary value is slightly below it.
    /// </summary>
    internal static double RoundForDisplay(double value, int digits) {
        double rounded = Math.Round(value, digits, MidpointRounding.AwayFromZero);
        double scaled = Math.Abs(value) * Math.Pow(10D, digits);
        double fraction = scaled - Math.Floor(scaled);
        if (Math.Abs(fraction - 0.5D) > 1E-6D || scaled > 1E15D) return rounded;
        if (!decimal.TryParse(value.ToString("G15", CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture,
                out decimal exact)) return rounded;
        return (double)Math.Round(exact, digits, MidpointRounding.AwayFromZero);
    }
}
