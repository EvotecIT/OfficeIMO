using System.Globalization;

namespace OfficeIMO.Project;

/// <summary>Strict value conversions for the explicitly supported Project expression profile.</summary>
internal readonly struct ProjectFormulaValue {
    internal readonly object Value;
    internal ProjectFormulaValue(object value) { Value = value; }
    internal decimal Number => Value switch {
        decimal number => number,
        bool flag => flag ? -1m : 0m,
        string text when decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal number) => number,
        _ => throw new InvalidDataException("The expression requires a numeric value.")
    };
    internal bool Flag => Value is bool flag ? flag : Number != 0m;
    internal decimal NumberIn(CultureInfo culture) => Value is string text ? decimal.Parse(text, NumberStyles.Float, culture) : Number;
    internal bool FlagIn(CultureInfo culture) => Value is bool flag ? flag : NumberIn(culture) != 0m;
    internal string TextIn(CultureInfo culture) => Value is decimal number ? number.ToString(culture) : Text;
    internal int IntegerIn(CultureInfo culture) {
        decimal number = NumberIn(culture);
        if (number != decimal.Truncate(number) || number < int.MinValue || number > int.MaxValue)
            throw new InvalidDataException("The expression requires an integral value within Int32 bounds.");
        return (int)number;
    }
    internal bool Logical => Value is bool flag ? flag : throw new NotSupportedException("Logical operators require Boolean operands; integer bitwise operations are outside this formula profile.");
    internal DateTime Date => Value is DateTime date ? date : throw new InvalidDataException("The expression requires a date value.");
    internal string Text => Value switch {
        string text => text,
        decimal number => number.ToString(CultureInfo.InvariantCulture),
        bool flag => flag ? "True" : "False",
        _ => throw new NotSupportedException("Date-to-text conversion requires a locale-specific format and is not implicit in this profile.")
    };
    internal static int Compare(ProjectFormulaValue left, ProjectFormulaValue right, CultureInfo? culture = null) {
        culture ??= CultureInfo.InvariantCulture;
        if (left.Value is DateTime || right.Value is DateTime) return left.Date.CompareTo(right.Date);
        if (left.Value is string && right.Value is string) return culture.CompareInfo.Compare(left.Text, right.Text, CompareOptions.IgnoreCase);
        return left.NumberIn(culture).CompareTo(right.NumberIn(culture));
    }
}
