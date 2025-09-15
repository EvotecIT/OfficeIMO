using System;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

internal static class CoerceValueHelper
{
    private const long DoublePrecisionSafeIntegerBound = 9_007_199_254_740_992L;
    private const ulong DoublePrecisionSafeIntegerBoundUnsigned = 9_007_199_254_740_992UL;
    private const int SharedStringCharacterLimit = 32_767;

    private static readonly CellValue EmptyStringTemplate = new(string.Empty);
    private static readonly CellValue TrueTemplate = new("1");
    private static readonly CellValue FalseTemplate = new("0");

    /// <summary>
    /// Converts an arbitrary CLR value into the <see cref="CellValue" /> and <see cref="CellValues" /> tuple used by Excel.
    /// </summary>
    /// <param name="value">The value to be represented in a worksheet cell.</param>
    /// <param name="sharedStringHandler">Delegate that materialises a <see cref="CellValue" /> entry for shared strings.</param>
    /// <returns>The OpenXML representation of the supplied value.</returns>
    /// <remarks>
    /// Integer values with an absolute magnitude above ±9,007,199,254,740,992 are written using their string form to prevent
    /// double precision loss while keeping the cell typed as <see cref="CellValues.Number" />.
    /// </remarks>
    /// <exception cref="ArgumentException">Thrown when the resulting shared-string payload exceeds Excel's 32,767 character limit.</exception>
    internal static (CellValue cellValue, CellValues type) Coerce(object? value, Func<string, CellValue> sharedStringHandler)
    {
        if (sharedStringHandler is null)
        {
            throw new ArgumentNullException(nameof(sharedStringHandler));
        }

        return value switch
        {
            null => HandleEmptyString(),
            System.DBNull => HandleEmptyString(),
            string s => HandleSharedString(s, sharedStringHandler, nameof(value)),
            double d => HandleNumber(d),
            float f => HandleNumber(Convert.ToDouble(f)),
            decimal dec => HandleDecimal(dec),
            int i => HandleSignedInteger(i),
            long l => HandleSignedInteger(l),
            DateTime dt => HandleNumber(dt.ToOADate()),
            DateTimeOffset dto => HandleNumber(dto.UtcDateTime.ToOADate()),
#if NET6_0_OR_GREATER
            DateOnly dateOnly => HandleNumber(dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate()),
            TimeOnly timeOnly => HandleNumber(timeOnly.ToTimeSpan().TotalDays),
#endif
            TimeSpan ts => HandleNumber(ts.TotalDays),
            bool b => HandleBoolean(b),
            uint ui => HandleUnsignedInteger(ui),
            ulong ul => HandleUnsignedInteger(ul),
            ushort us => HandleUnsignedInteger(us),
            byte by => HandleUnsignedInteger(by),
            sbyte sb => HandleSignedInteger(sb),
            short sh => HandleSignedInteger(sh),
            Guid guid => HandleSharedString(guid.ToString(), sharedStringHandler, nameof(value)),
            Enum e => HandleSharedString(e.ToString(), sharedStringHandler, nameof(value)),
            char ch => HandleSharedString(ch.ToString(), sharedStringHandler, nameof(value)),
            Uri uri => HandleSharedString(uri.ToString(), sharedStringHandler, nameof(value)),
            _ => HandleSharedString(value.ToString() ?? string.Empty, sharedStringHandler, nameof(value))
        };
    }

    internal static (CellValue, CellValues) HandleNumber(double number) =>
        (CreateNumberCellValue(number), CellValues.Number);

    internal static (CellValue, CellValues) HandleDecimal(decimal value) =>
        (CreateTextCellValue(value.ToString(CultureInfo.InvariantCulture)), CellValues.Number);

    internal static (CellValue, CellValues) HandleSignedInteger(long integer) =>
        integer is >= -DoublePrecisionSafeIntegerBound and <= DoublePrecisionSafeIntegerBound
            ? HandleNumber(integer)
            : (CreateTextCellValue(integer.ToString(CultureInfo.InvariantCulture)), CellValues.Number);

    internal static (CellValue, CellValues) HandleUnsignedInteger(ulong integer) =>
        integer <= DoublePrecisionSafeIntegerBoundUnsigned
            ? HandleNumber(integer)
            : (CreateTextCellValue(integer.ToString(CultureInfo.InvariantCulture)), CellValues.Number);

    internal static (CellValue, CellValues) HandleBoolean(bool value) =>
        (CloneCellValue(value ? TrueTemplate : FalseTemplate), CellValues.Boolean);

    internal static (CellValue, CellValues) HandleEmptyString() =>
        (CloneCellValue(EmptyStringTemplate), CellValues.String);

    internal static (CellValue, CellValues) HandleSharedString(string text, Func<string, CellValue> sharedStringHandler, string? paramName = null)
    {
        if (sharedStringHandler is null)
        {
            throw new ArgumentNullException(nameof(sharedStringHandler));
        }

        ValidateSharedStringLength(text, paramName ?? nameof(text));
        return (sharedStringHandler(text), CellValues.SharedString);
    }

    internal static void ValidateSharedStringLength(string text, string paramName)
    {
        if (text.Length > SharedStringCharacterLimit)
        {
            throw new ArgumentException("String exceeds Excel's limit of 32,767 characters.", paramName);
        }
    }

    private static CellValue CreateNumberCellValue(double number) =>
        new(number.ToString(CultureInfo.InvariantCulture));

    private static CellValue CreateTextCellValue(string text) => new(text);

    private static CellValue CloneCellValue(CellValue template) => new(template.Text);
}
