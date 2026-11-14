using System.Buffers;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    /// <summary>Growable pooled UTF-16 or UTF-8 units for completed data-reader rows.</summary>
    internal sealed class RecordBuffer<TChar> : IDisposable
        where TChar : unmanaged
    {
        private TChar[] _units = ArrayPool<TChar>.Shared.Rent(DataReaderFlushThreshold + 4096);

        internal int Length { get; set; }

        internal ReadOnlySpan<TChar> Written => _units.AsSpan(0, Length);

        /// <summary>The first units, which stay readable after <see cref="Length"/> is reset.</summary>
        internal ReadOnlySpan<TChar> Prefix(int length) => _units.AsSpan(0, length);

#if !NET8_0_OR_GREATER
        /// <summary>The backing array, for writers that take arrays rather than spans.</summary>
        internal TChar[] Units => _units;
#endif

        internal Span<TChar> GetSpan(int sizeHint)
        {
            if (_units.Length - Length < sizeHint)
            {
                TChar[] larger = ArrayPool<TChar>.Shared.Rent(Math.Max(_units.Length * 2, Length + sizeHint));
                _units.AsSpan(0, Length).CopyTo(larger);
                ArrayPool<TChar>.Shared.Return(_units);
                _units = larger;
            }

            return _units.AsSpan(Length);
        }

        internal void Write(TChar value)
        {
            GetSpan(1)[0] = value;
            Length++;
        }

        internal void Write(ReadOnlySpan<TChar> value)
        {
            value.CopyTo(GetSpan(value.Length));
            Length += value.Length;
        }

        public void Dispose()
        {
            ArrayPool<TChar>.Shared.Return(_units, clearArray: true);
            _units = Array.Empty<TChar>();
        }
    }

    /// <summary>
    /// Appends one typed data-reader record. Strings, integers, doubles, decimals, booleans and dates are
    /// formatted straight into the buffer; every other value goes through the row appenders of the general
    /// typed path, so the text is the same as that path writes.
    /// </summary>
    internal static void AppendDataReaderRecord(
        RecordBuffer<char> output,
        StringBuilder scratch,
        IDataRecord reader,
        DataReaderFieldKind[] fieldKinds,
        char delimiter,
        string newLine,
        CultureInfo culture,
        bool formatted,
        string? dateTimeFormat,
        bool useUtc,
        string? nullValue,
        bool boxedValues = false)
    {
        bool invariantComma = delimiter == ',' && ReferenceEquals(culture, CultureInfo.InvariantCulture);
        string? format = string.IsNullOrEmpty(dateTimeFormat) ? null : dateTimeFormat;
        for (var i = 0; i < fieldKinds.Length; i++)
        {
            if (i > 0)
            {
                output.Write(delimiter);
            }

            if (TryAppendRecordValue(output, reader, i, fieldKinds[i], delimiter, culture, invariantComma, formatted, format, useUtc, nullValue, boxedValues))
            {
                continue;
            }

#if NET8_0_OR_GREATER
            scratch.Clear();
            if (formatted)
            {
                if (fieldKinds[i] == DataReaderFieldKind.Object || !TryAppendFormattedTypedDataReaderValue(
                        scratch, reader, i, fieldKinds[i], delimiter, culture, format, useUtc, nullValue))
                {
                    scratch.Clear();
                    fieldKinds[i] = DataReaderFieldKind.Object;
                    AppendFormattedDataReaderObjectValue(scratch, reader.GetValue(i), delimiter, culture, dateTimeFormat, useUtc, nullValue);
                }
            }
            else if (fieldKinds[i] == DataReaderFieldKind.Object || !TryAppendTypedDataReaderValue(scratch, reader, i, fieldKinds[i], delimiter, culture))
            {
                scratch.Clear();
                fieldKinds[i] = DataReaderFieldKind.Object;
                AppendDataReaderObjectValue(scratch, reader.GetValue(i), delimiter, culture);
            }

            foreach (ReadOnlyMemory<char> chunk in scratch.GetChunks())
            {
                output.Write(chunk.Span);
            }
#else
            // Other kinds, and a column whose typed read was rejected, take the boxed value from here on.
            fieldKinds[i] = DataReaderFieldKind.Object;
            AppendRecordObjectValue(output, reader.GetValue(i), delimiter, culture, dateTimeFormat, useUtc, nullValue);
#endif
        }

        output.Write(newLine.AsSpan());
    }

    /// <summary>Writes the common kinds directly; false leaves the value to the general appenders.</summary>
    private static bool TryAppendRecordValue(
        RecordBuffer<char> output,
        IDataRecord reader,
        int ordinal,
        DataReaderFieldKind fieldKind,
        char delimiter,
        CultureInfo culture,
        bool invariantComma,
        bool formatted,
        string? format,
        bool useUtc,
        string? nullValue,
        bool boxedValues)
    {
        switch (fieldKind)
        {
            case DataReaderFieldKind.String:
            case DataReaderFieldKind.Int32:
            case DataReaderFieldKind.Int64:
            case DataReaderFieldKind.Double:
            case DataReaderFieldKind.Decimal:
            case DataReaderFieldKind.Boolean:
            case DataReaderFieldKind.DateTime:
                break;
            default:
                return false;
        }

        try
        {
            // A data table reader looks a row up on every call; one boxed read answers both the null check
            // and the value, and a table column holds only its declared type or DBNull.
            object? boxed = null;
            if (boxedValues ? (boxed = reader.GetValue(ordinal)) is DBNull : reader.IsDBNull(ordinal))
            {
                // Without formatting options there is no null marker, and an absent one writes nothing.
                if (formatted && nullValue is not null)
                {
                    AppendEscapedRecordText(output, nullValue.AsSpan(), delimiter);
                }

                return true;
            }

            switch (fieldKind)
            {
                case DataReaderFieldKind.String:
                    AppendEscapedRecordText(output, (boxedValues ? (string)boxed! : reader.GetString(ordinal)).AsSpan(), delimiter);
                    return true;
                case DataReaderFieldKind.Int32:
                    return AppendRecordScalar(output, boxedValues ? (int)boxed! : reader.GetInt32(ordinal), delimiter, culture, invariantComma);
                case DataReaderFieldKind.Int64:
                    return AppendRecordScalar(output, boxedValues ? (long)boxed! : reader.GetInt64(ordinal), delimiter, culture, invariantComma);
                case DataReaderFieldKind.Double:
                    return AppendRecordScalar(output, boxedValues ? (double)boxed! : reader.GetDouble(ordinal), delimiter, culture, invariantComma);
                case DataReaderFieldKind.Decimal:
                    return AppendRecordScalar(output, boxedValues ? (decimal)boxed! : reader.GetDecimal(ordinal), delimiter, culture, invariantComma);
                case DataReaderFieldKind.Boolean:
                    AppendEscapedRecordText(output, ((boxedValues ? (bool)boxed! : reader.GetBoolean(ordinal)) ? "True" : "False").AsSpan(), delimiter);
                    return true;
                default:
                    var value = boxedValues ? (DateTime)boxed! : reader.GetDateTime(ordinal);
                    if (!formatted)
                    {
                        return AppendRecordScalar(output, value, delimiter, culture, invariantComma);
                    }

                    if (useUtc)
                    {
                        value = value.ToUniversalTime();
                    }

                    // Dates are formatted in place and moved aside only when they need quotes.
                    Span<char> destination = output.GetSpan(64).Slice(0, 64);
                    if (!TryFormatDate(value, destination, out var written, format, culture))
                    {
                        return false;
                    }

                    if (!NeedsQuotes(destination.Slice(0, written), delimiter))
                    {
                        output.Length += written;
                        return true;
                    }

                    Span<char> characters = stackalloc char[64];
                    destination.Slice(0, written).CopyTo(characters);
                    AppendEscapedRecordText(output, characters.Slice(0, written), delimiter);
                    return true;
            }
        }
        catch (InvalidCastException)
        {
            return false;
        }
        catch (NotSupportedException)
        {
            return false;
        }
        catch (NotImplementedException)
        {
            return false;
        }
    }

#if NET8_0_OR_GREATER
    /// <summary>Formats a known scalar as the general path does: unescaped for invariant comma output, escaped otherwise.</summary>
    private static bool AppendRecordScalar<T>(RecordBuffer<char> output, T value, char delimiter, CultureInfo culture, bool invariantComma)
        where T : ISpanFormattable, IUtf8SpanFormattable
    {
        if (invariantComma)
        {
            if (!TryFormatInvariant(value, output.GetSpan(64), out var written))
            {
                return false;
            }

            output.Length += written;
            return true;
        }

        Span<char> characters = stackalloc char[128];
        if (!value.TryFormat(characters, out var length, default, culture))
        {
            return false;
        }

        AppendEscapedRecordText(output, characters[..length], delimiter);
        return true;
    }
#else
    // .NET Framework has no span formatting: integers, decimals and invariant dates are written here as
    // ToString(null, culture) writes them, straight into the buffer, and other values take the runtime's string.

    private static bool AppendRecordScalar(RecordBuffer<char> output, int value, char delimiter, CultureInfo culture, bool invariantComma) =>
        AppendRecordDecimal(output, value < 0 ? (ulong)-(long)value : (ulong)value, value < 0, 0, delimiter, culture, invariantComma);

    private static bool AppendRecordScalar(RecordBuffer<char> output, long value, char delimiter, CultureInfo culture, bool invariantComma) =>
        AppendRecordDecimal(output, value < 0 ? 0 - (ulong)value : (ulong)value, value < 0, 0, delimiter, culture, invariantComma);

    /// <summary>
    /// Writes a decimal as ToString(null, culture) does: the culture's negative sign, the coefficient's digits and
    /// every digit of the scale after the culture's decimal separator. Zero and coefficients past 64 bits take the
    /// runtime's text.
    /// </summary>
    private static bool AppendRecordScalar(RecordBuffer<char> output, decimal value, char delimiter, CultureInfo culture, bool invariantComma)
    {
        var parts = new DecimalParts { Value = value };
        ulong coefficient = ((ulong)parts.Mid << 32) | parts.Low;
        if (!BitConverter.IsLittleEndian || parts.High != 0 || coefficient == 0)
        {
            AppendScalarText(output, value.ToString(culture).AsSpan(), delimiter, invariantComma);
            return true;
        }

        return AppendRecordDecimal(output, coefficient, parts.Flags < 0, (parts.Flags >> 16) & 0xFF, delimiter, culture, invariantComma);
    }

    /// <summary>
    /// Writes coefficient * 10^-scale with the culture's negative sign and decimal separator, every fraction digit
    /// kept. Invariant comma output goes straight into the buffer; other output is escaped like any text.
    /// </summary>
    private static bool AppendRecordDecimal(
        RecordBuffer<char> output, ulong coefficient, bool negative, int scale, char delimiter, CultureInfo culture, bool invariantComma)
    {
        if (invariantComma)
        {
            Span<char> destination = output.GetSpan(MaximumScalarLength);
            int start = 0;
            if (negative)
            {
                destination[start++] = '-';
            }

            output.Length += start + WriteFixed(coefficient, scale, '.', destination.Slice(start));
            return true;
        }

        NumberFormatInfo number = culture.NumberFormat;
        string sign = negative ? number.NegativeSign : "-";
        string separator = scale > 0 ? number.NumberDecimalSeparator : ".";
        if (sign.Length != 1 || separator.Length != 1)
        {
            // Signs and separators of other lengths are rare; the general path writes them.
            return false;
        }

        Span<char> characters = stackalloc char[MaximumScalarLength];
        int position = 0;
        if (negative)
        {
            characters[position++] = sign[0];
        }

        position += WriteFixed(coefficient, scale, separator[0], characters.Slice(position));
        AppendEscapedRecordText(output, characters.Slice(0, position), delimiter);
        return true;
    }

    /// <summary>The longest integer or decimal <see cref="AppendRecordDecimal"/> writes: a sign, a separator and 29 digits after a zero.</summary>
    private const int MaximumScalarLength = 48;

    /// <summary>Writes digits * 10^-scale, every fraction digit kept, as "0.0dd" below one, and returns the length.</summary>
    private static int WriteFixed(ulong digits, int scale, char separator, Span<char> destination)
    {
        int count = CountDigits(digits);
        int length = count > scale ? count + (scale > 0 ? 1 : 0) : scale + 2;
        int position = length;
        for (int fraction = 0; fraction < scale; fraction++)
        {
            ulong quotient = digits / 10;
            destination[--position] = (char)('0' + (int)(digits - quotient * 10));
            digits = quotient;
        }

        if (scale > 0)
        {
            destination[--position] = separator;
        }

        if (count <= scale)
        {
            destination[0] = '0';
            return length;
        }

        // The integer part usually fits in 32 bits, where division is cheaper.
        while (digits > uint.MaxValue)
        {
            ulong quotient = digits / 10;
            destination[--position] = (char)('0' + (int)(digits - quotient * 10));
            digits = quotient;
        }

        uint small = (uint)digits;
        while (small >= 10)
        {
            uint quotient = small / 10;
            destination[--position] = (char)('0' + (int)(small - quotient * 10));
            small = quotient;
        }

        destination[--position] = (char)('0' + (int)small);
        return length;
    }

    private static int CountDigits(ulong value)
    {
        int count = 1;
        while (value >= 10_000)
        {
            value /= 10_000;
            count += 4;
        }

        return value >= 1_000 ? count + 3 : value >= 100 ? count + 2 : value >= 10 ? count + 1 : count;
    }

    private static bool AppendRecordScalar(RecordBuffer<char> output, double value, char delimiter, CultureInfo culture, bool invariantComma)
    {
        AppendScalarText(output, value.ToString(culture).AsSpan(), delimiter, invariantComma);
        return true;
    }

    private static bool AppendRecordScalar(RecordBuffer<char> output, DateTime value, char delimiter, CultureInfo culture, bool invariantComma)
    {
        if (ReferenceEquals(culture, CultureInfo.InvariantCulture))
        {
            // The invariant general date text has no comma, quote or line break.
            Span<char> characters = stackalloc char[32];
            CsvDateTemplate.TryFormatInvariantGeneral(value, characters, out int written);
            AppendScalarText(output, characters.Slice(0, written), delimiter, invariantComma);
            return true;
        }

        AppendEscapedRecordText(output, value.ToString(culture).AsSpan(), delimiter);
        return true;
    }

    /// <summary>Invariant comma output never quotes a scalar; anything else applies the default quoting rule.</summary>
    private static void AppendScalarText(RecordBuffer<char> output, ReadOnlySpan<char> text, char delimiter, bool invariantComma)
    {
        if (invariantComma)
        {
            output.Write(text);
        }
        else
        {
            AppendEscapedRecordText(output, text, delimiter);
        }
    }

    /// <summary>The fields of a decimal, which keeps its flags, high, low and middle 32 bits in that order on little-endian machines.</summary>
    [StructLayout(LayoutKind.Explicit)]
    private struct DecimalParts
    {
        [FieldOffset(0)] public decimal Value;
        [FieldOffset(0)] public int Flags;
        [FieldOffset(4)] public uint High;
        [FieldOffset(8)] public uint Low;
        [FieldOffset(12)] public uint Mid;
    }

    private static void AppendRecordObjectValue(
        RecordBuffer<char> output, object? value, char delimiter, CultureInfo culture, string? dateTimeFormat, bool useUtc, string? nullValue)
    {
        if (value is null || value is DBNull)
        {
            if (nullValue is not null)
            {
                AppendEscapedRecordText(output, nullValue.AsSpan(), delimiter);
            }

            return;
        }

        string text = value as string ?? FormatValue(value, culture, dateTimeFormat, useUtc, nullValue);
        AppendEscapedRecordText(output, text.AsSpan(), delimiter);
    }
#endif

    /// <summary>Applies the default quoting rule: quote text with a delimiter, quote or line break and double its quotes.</summary>
    private static void AppendEscapedRecordText(RecordBuffer<char> output, ReadOnlySpan<char> text, char delimiter)
    {
        if (!NeedsQuotes(text, delimiter))
        {
            output.Write(text);
            return;
        }

        output.Write('"');
        int quote = text.IndexOf('"');
        while (quote >= 0)
        {
            output.Write(text.Slice(0, quote + 1));
            output.Write('"');
            text = text.Slice(quote + 1);
            quote = text.IndexOf('"');
        }

        output.Write(text);
        output.Write('"');
    }
}
