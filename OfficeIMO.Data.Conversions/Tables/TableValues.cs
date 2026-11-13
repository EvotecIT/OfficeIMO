namespace OfficeIMO.Data.Conversions;

/// <summary>Turns cell text and kind into the CLR values a data reader returns.</summary>
internal static class TableValues {
    public static bool FitsInt64(ReadOnlySpan<char> text) =>
        long.TryParse(text.ToString(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out _);

    /// <summary>Whether the decimal type holds the number without changing its value.</summary>
    public static bool FitsDecimal(ReadOnlySpan<char> text) {
        string value = text.ToString();
        if (!decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal number)) return false;
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double binary)) return false;
        return (double)number == binary && (number != 0m || binary == 0d);
    }

    public static Type ColumnType(TableColumnNode column) {
        if (column.JoinsArray || column.MixedKinds || column.Kind is not { } kind) return column.JoinsArray ? typeof(string) : typeof(object);
        return kind switch {
            DataScalarKind.Boolean => typeof(bool),
            DataScalarKind.Integer => column.IntegerOverflow ? typeof(decimal) : typeof(long),
            DataScalarKind.Float => column.FloatNotDecimal ? typeof(double) : typeof(decimal),
            DataScalarKind.String => typeof(string),
            DataScalarKind.OffsetDateTime => typeof(DateTimeOffset),
            DataScalarKind.LocalDateTime or DataScalarKind.LocalDate => typeof(DateTime),
            DataScalarKind.LocalTime => typeof(TimeSpan),
            _ => typeof(object)
        };
    }

    /// <summary>Converts one cell to the column's type, or to its own natural type in an untyped column.</summary>
    public static object Convert(DataScalarKind kind, string text, Type column) {
        switch (kind) {
            case DataScalarKind.Null:
                return DBNull.Value;
            case DataScalarKind.Boolean:
                return text.Length == 4;
            case DataScalarKind.Integer:
                if (column != typeof(decimal) && column != typeof(double)
                    && long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long integer)) return integer;
                goto case DataScalarKind.Float;
            case DataScalarKind.Float:
                if (DataScalarKinds.IsNonFinite(text.AsSpan())) {
                    return text[0] == 'N' ? double.NaN : text[0] == '-' ? double.NegativeInfinity : double.PositiveInfinity;
                }

                if (column != typeof(double) && FitsDecimal(text.AsSpan())) return decimal.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture);
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return number;
                return text;
            case DataScalarKind.OffsetDateTime:
                return DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTimeOffset offset) ? offset : text;
            case DataScalarKind.LocalDateTime:
            case DataScalarKind.LocalDate:
                return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date) ? date : text;
            case DataScalarKind.LocalTime:
                return TimeSpan.TryParse(text, CultureInfo.InvariantCulture, out TimeSpan time) ? time : text;
            default:
                return text;
        }
    }

    /// <summary>Formats a cell as CSV text: numbers in the culture, everything else in its canonical spelling.</summary>
    public static string Text(DataScalarKind kind, string text, CultureInfo culture) {
        if ((kind == DataScalarKind.Integer || kind == DataScalarKind.Float) && !ReferenceEquals(culture, CultureInfo.InvariantCulture)
            && culture.NumberFormat.NumberDecimalSeparator != "." && !DataScalarKinds.IsNonFinite(text.AsSpan())) {
            return text.Replace(".", culture.NumberFormat.NumberDecimalSeparator);
        }

        return text;
    }
}
