using System;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

internal static class CoerceValueHelper
{
    private const long DoublePrecisionSafeIntegerBound = 9_007_199_254_740_992L;
    private const ulong DoublePrecisionSafeIntegerBoundUnsigned = 9_007_199_254_740_992UL;

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
        return value switch
        {
            null => (new CellValue(string.Empty), CellValues.String),
            System.DBNull => (new CellValue(string.Empty), CellValues.String),
            string s => HandleSharedString(s),
            double d => HandleNumber(d),
            float f => HandleNumber(Convert.ToDouble(f)),
            decimal dec => (new CellValue(dec.ToString(CultureInfo.InvariantCulture)), CellValues.Number),
            int i => HandleSignedInteger(i),
            long l => HandleSignedInteger(l),
            DateTime dt => HandleNumber(dt.ToOADate()),
            DateTimeOffset dto => HandleNumber(dto.UtcDateTime.ToOADate()),
#if NET6_0_OR_GREATER
            DateOnly dateOnly => HandleNumber(dateOnly.ToDateTime(TimeOnly.MinValue).ToOADate()),
            TimeOnly timeOnly => HandleNumber(timeOnly.ToTimeSpan().TotalDays),
#endif
            TimeSpan ts => HandleNumber(ts.TotalDays),
            bool b => (new CellValue(b ? "1" : "0"), CellValues.Boolean),
            uint ui => HandleUnsignedInteger(ui),
            ulong ul => HandleUnsignedInteger(ul),
            ushort us => HandleUnsignedInteger(us),
            byte by => HandleUnsignedInteger(by),
            sbyte sb => HandleSignedInteger(sb),
            short sh => HandleSignedInteger(sh),
            Guid guid => HandleSharedString(guid.ToString()),
            Enum e => HandleSharedString(e.ToString()),
            char ch => HandleSharedString(ch.ToString()),
            Uri uri => HandleSharedString(uri.ToString()),
            _ => HandleSharedString(value.ToString() ?? string.Empty)
        };

        (CellValue, CellValues) HandleNumber(double number) =>
            (new CellValue(number.ToString(CultureInfo.InvariantCulture)), CellValues.Number);

        (CellValue, CellValues) HandleSignedInteger(long integer) =>
            integer is >= -DoublePrecisionSafeIntegerBound and <= DoublePrecisionSafeIntegerBound
                ? HandleNumber(integer)
                : (new CellValue(integer.ToString(CultureInfo.InvariantCulture)), CellValues.Number);

        (CellValue, CellValues) HandleUnsignedInteger(ulong integer) =>
            integer <= DoublePrecisionSafeIntegerBoundUnsigned
                ? HandleNumber(integer)
                : (new CellValue(integer.ToString(CultureInfo.InvariantCulture)), CellValues.Number);

        (CellValue, CellValues) HandleSharedString(string text)
        {
            ValidateSharedStringLength(text);
            return (sharedStringHandler(text), CellValues.SharedString);
        }

        void ValidateSharedStringLength(string text)
        {
            if (text.Length > 32_767)
            {
                throw new ArgumentException("String exceeds Excel's limit of 32,767 characters.", nameof(value));
            }
        }
    }
}
