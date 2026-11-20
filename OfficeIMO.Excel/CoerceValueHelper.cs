using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

internal static class CoerceValueHelper {
    internal const int SharedStringCharacterLimit = 32_767;

    private static readonly CellValue EmptyStringTemplate = new(string.Empty);
    private static readonly CellValue TrueTemplate = new("1");
    private static readonly CellValue FalseTemplate = new("0");
    /// <summary>
    /// Marker strategy meaning "write DateTimeOffset values as ISO 8601 text". Write paths compare by reference, so
    /// the conversion itself only runs if a path has no text handling and then keeps the previous local-time serial.
    /// </summary>
    internal static readonly Func<DateTimeOffset, DateTime> DateTimeOffsetTextStrategy = static dto => dto.LocalDateTime;
    // Excel cannot represent serial dates earlier than 1900-01-01; 0 = 1899-12-30 and 2 = 1900-01-01.
    private const double ExcelMinimumSupportedSerial = 2d;
    private static readonly DateTime ExcelMinimumSupportedDate = DateTime.FromOADate(ExcelMinimumSupportedSerial);

    /// <summary>
    /// Converts an arbitrary CLR value into the <see cref="CellValue" /> and <see cref="CellValues" /> tuple used by Excel.
    /// </summary>
    /// <param name="value">The value to be represented in a worksheet cell.</param>
    /// <param name="sharedStringHandler">Delegate that materialises a <see cref="CellValue" /> entry for shared strings.</param>
    /// <param name="dateTimeOffsetStrategy">
    /// Optional delegate used to convert <see cref="DateTimeOffset"/> instances into <see cref="DateTime"/> values before
    /// calculating the Excel serial number. When omitted or <see cref="DateTimeOffsetTextStrategy"/>, the value is written
    /// as ISO 8601 text with its offset. The provided strategy only influences the numeric serial value that is written;
    /// cell formatting must still be applied separately.
    /// </param>
    /// <param name="dateSystem">Workbook date system used when serializing <see cref="DateTime"/> and <see cref="DateTimeOffset"/> values.</param>
    /// <returns>The OpenXML representation of the supplied value.</returns>
    /// <remarks>
    /// Decimal and 64-bit integer values with more than 15 significant digits are written as text, because Excel and
    /// every reader hold numeric cells as doubles. Byte arrays are written as base64 text.
    /// </remarks>
    /// <exception cref="ArgumentException">Thrown when the resulting shared-string payload exceeds Excel's 32,767 character limit.</exception>
    internal static (CellValue cellValue, CellValues type) Coerce(
        object? value,
        Func<string, CellValue> sharedStringHandler,
        Func<DateTimeOffset, DateTime>? dateTimeOffsetStrategy = null,
        ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred) {
        ArgumentNullException.ThrowIfNull(sharedStringHandler);

        dateTimeOffsetStrategy ??= DateTimeOffsetTextStrategy;

        return value switch {
            null => HandleEmptyString(),
            System.DBNull => HandleEmptyString(),
            string s => HandleSharedString(s, sharedStringHandler, nameof(value)),
            double d => HandleNumber(d),
            float f => HandleNumber((double)f),
            decimal dec when ExcelExactNumber.FitsNumberCell(dec) => HandleDecimal(dec),
            decimal dec => HandleSharedString(dec.ToString(CultureInfo.InvariantCulture), sharedStringHandler, nameof(value)),
            int i => HandleSignedInteger(i),
            long l when ExcelExactNumber.FitsNumberCell(l) => HandleSignedInteger(l),
            long l => HandleSharedString(InvariantNumberText.Get(l), sharedStringHandler, nameof(value)),
            DateTime dt => HandleNumber(ExcelDateSystemConverter.ToSerial(dt, dateSystem)),
            DateTimeOffset dto => HandleDateTimeOffset(dto, sharedStringHandler, dateTimeOffsetStrategy, dateSystem, nameof(value)),
#if NET6_0_OR_GREATER
            DateOnly dateOnly => HandleNumber(ExcelDateSystemConverter.ToSerial(dateOnly.ToDateTime(TimeOnly.MinValue), dateSystem)),
            TimeOnly timeOnly => HandleNumber(ExcelDateSystemConverter.ToSerial(timeOnly.ToTimeSpan())),
#endif
            TimeSpan ts => HandleNumber(ExcelDateSystemConverter.ToSerial(ts)),
            bool b => HandleBoolean(b),
            uint ui => HandleUnsignedInteger(ui),
            ulong ul when ExcelExactNumber.FitsNumberCell(ul) => HandleUnsignedInteger(ul),
            ulong ul => HandleSharedString(InvariantNumberText.Get(ul), sharedStringHandler, nameof(value)),
            ushort us => HandleUnsignedInteger(us),
            byte by => HandleUnsignedInteger(by),
            sbyte sb => HandleSignedInteger(sb),
            short sh => HandleSignedInteger(sh),
            Guid guid => HandleSharedString(guid.ToString(), sharedStringHandler, nameof(value)),
            Enum e => HandleSharedString(e.ToString(), sharedStringHandler, nameof(value)),
            char ch => HandleSharedString(ch.ToString(), sharedStringHandler, nameof(value)),
            Uri uri => HandleSharedString(uri.ToString(), sharedStringHandler, nameof(value)),
            byte[] bytes => HandleSharedString(System.Convert.ToBase64String(bytes), sharedStringHandler, nameof(value)),
            _ => HandleSharedString(value.ToString() ?? string.Empty, sharedStringHandler, nameof(value))
        };
    }

    /// <summary>Excel has no NaN or infinite cell value; such numbers are stored as the #NUM! error Excel shows for them.</summary>
    internal static (CellValue, CellValues) HandleNumber(double number) =>
        double.IsNaN(number) || double.IsInfinity(number)
            ? (CreateTextCellValue(ExcelSheet.NonFiniteNumberErrorText), CellValues.Error)
            : (CreateNumberCellValue(number), CellValues.Number);

    internal static (CellValue, CellValues) HandleDecimal(decimal value) =>
        (CreateTextCellValue(value.ToString(CultureInfo.InvariantCulture)), CellValues.Number);

    internal static (CellValue, CellValues) HandleSignedInteger(long integer) =>
        (CreateTextCellValue(InvariantNumberText.Get(integer)), CellValues.Number);

    internal static (CellValue, CellValues) HandleUnsignedInteger(ulong integer) =>
        (CreateTextCellValue(InvariantNumberText.Get(integer)), CellValues.Number);

    internal static (CellValue, CellValues) HandleDateTimeOffset(
        DateTimeOffset value,
        Func<string, CellValue> sharedStringHandler,
        Func<DateTimeOffset, DateTime> dateTimeOffsetStrategy,
        ExcelDateSystem dateSystem,
        string? paramName) {
        ArgumentNullException.ThrowIfNull(sharedStringHandler);

        ArgumentNullException.ThrowIfNull(dateTimeOffsetStrategy);

        return TryConvertDateTimeOffset(
                value,
                dateTimeOffsetStrategy,
                dateSystem,
                out _,
                out double serial)
            ? HandleNumber(serial)
            : HandleDateTimeOffsetFallback(value, sharedStringHandler, paramName);
    }

    internal static bool TryConvertDateTimeOffset(
        DateTimeOffset value,
        Func<DateTimeOffset, DateTime> dateTimeOffsetStrategy,
        ExcelDateSystem dateSystem,
        out DateTime converted,
        out double serial) {
        if (IsTextStrategy(dateTimeOffsetStrategy)) {
            converted = default;
            serial = 0D;
            return false;
        }

        try {
            converted = dateTimeOffsetStrategy(value);
        } catch (Exception ex) {
            throw new InvalidOperationException("The configured DateTimeOffset write strategy threw an exception.", ex);
        }

        if (value.UtcDateTime < ExcelMinimumSupportedDate) {
            serial = 0D;
            return false;
        }
        try {
            serial = ExcelDateSystemConverter.ToSerial(converted, dateSystem);
            return true;
        } catch (ArgumentException) {
            serial = 0D;
            return false;
        } catch (OverflowException) {
            serial = 0D;
            return false;
        }
    }

    internal static bool IsTextStrategy(Func<DateTimeOffset, DateTime> dateTimeOffsetStrategy) =>
        ReferenceEquals(dateTimeOffsetStrategy, DateTimeOffsetTextStrategy);

    /// <summary>The text written for a DateTimeOffset kept as text: ISO 8601 with every fractional digit and the offset.</summary>
    internal static string FormatDateTimeOffsetText(DateTimeOffset value) => OfficeRoundTripText.FormatRoundTripPattern(value);

    /// <summary>Returns the exact text for a decimal or 64-bit integer that has more than 15 significant digits.</summary>
    internal static bool TryGetWideNumberText(object? value, out string text) {
        switch (value) {
            case decimal number when !ExcelExactNumber.FitsNumberCell(number):
                text = number.ToString(CultureInfo.InvariantCulture);
                return true;
            case long number when !ExcelExactNumber.FitsNumberCell(number):
                text = InvariantNumberText.Get(number);
                return true;
            case ulong number when !ExcelExactNumber.FitsNumberCell(number):
                text = InvariantNumberText.Get(number);
                return true;
            default:
                text = string.Empty;
                return false;
        }
    }

    private static (CellValue, CellValues) HandleDateTimeOffsetFallback(
        DateTimeOffset value,
        Func<string, CellValue> sharedStringHandler,
        string? paramName) =>
        HandleSharedString(FormatDateTimeOffsetText(value), sharedStringHandler, paramName ?? nameof(value));

    internal static (CellValue, CellValues) HandleBoolean(bool value) =>
        (CloneCellValue(value ? TrueTemplate : FalseTemplate), CellValues.Boolean);

    internal static (CellValue, CellValues) HandleEmptyString() =>
        (CloneCellValue(EmptyStringTemplate), CellValues.String);

    internal static (CellValue, CellValues) HandleSharedString(string text, Func<string, CellValue> sharedStringHandler, string? paramName = null) {
        ArgumentNullException.ThrowIfNull(sharedStringHandler);

        ValidateSharedStringLength(text, paramName ?? nameof(text));
        return (sharedStringHandler(text), CellValues.SharedString);
    }

    internal static void ValidateSharedStringLength(string text, string paramName) {
        if (text.Length > SharedStringCharacterLimit) {
            throw new ArgumentException("String exceeds Excel's limit of 32,767 characters.", paramName);
        }
    }

    private static CellValue CreateNumberCellValue(double number) =>
        new(InvariantNumberText.Get(number));

    private static CellValue CreateTextCellValue(string text) => new(text);

    private static CellValue CloneCellValue(CellValue template) => new(template.Text);
}
