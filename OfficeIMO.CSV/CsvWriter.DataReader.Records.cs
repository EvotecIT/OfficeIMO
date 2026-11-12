#if NET8_0_OR_GREATER
using System.Buffers;
using System.Data;
using System.Globalization;
using System.Runtime.CompilerServices;
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
        }

        output.Write(newLine);
    }

    /// <summary>True when every field is a kind <see cref="AppendInvariantRecord"/> writes.</summary>
    internal static bool HasOnlyInvariantRecordKinds(DataReaderFieldKind[] fieldKinds)
    {
        foreach (DataReaderFieldKind kind in fieldKinds)
        {
            if (kind is not (DataReaderFieldKind.String or DataReaderFieldKind.Int32 or DataReaderFieldKind.Int64
                or DataReaderFieldKind.Double or DataReaderFieldKind.Decimal or DataReaderFieldKind.Boolean or DataReaderFieldKind.DateTime))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Appends a record of strings, integers, doubles, decimals, booleans and dates for comma output in the
    /// invariant culture without value formatting options, as <see cref="AppendDataReaderRecord"/> writes it,
    /// in one pass with no per-field exception handling. A reader that rejects a typed read throws out of here;
    /// the caller then rewinds the row and writes it through <see cref="AppendDataReaderRecord"/>.
    /// </summary>
    internal static void AppendInvariantRecord<TFields, TChar>(
        RecordBuffer<TChar> output, ref TFields fields, DataReaderFieldKind[] fieldKinds, ReadOnlySpan<TChar> newLine, Encoding? encoding)
        where TFields : struct, IRecordFields
        where TChar : unmanaged
    {
        // Fields are written through a local span of the free buffer, committed when it runs short and at the end;
        // every check leaves room for a delimiter and the longest scalar.
        Span<TChar> free = output.GetSpan(InvariantFieldSpace);
        int position = 0;
        for (var i = 0; i < fieldKinds.Length; i++)
        {
            if (free.Length - position < InvariantFieldSpace)
            {
                output.Length += position;
                position = 0;
                free = output.GetSpan(InvariantFieldSpace);
            }

            if (i > 0)
            {
                free[position++] = CsvShortestDouble.Unit<TChar>(',');
            }

            if (fields.IsNull(i))
            {
                continue;
            }

            Span<TChar> destination = free.Slice(position);
            switch (fieldKinds[i])
            {
                case DataReaderFieldKind.String:
                    string text = fields.GetString(i);
                    if (TryCopyPlainText(text, destination, out bool quoted))
                    {
                        position += text.Length;
                        break;
                    }

                    output.Length += position;
                    position = 0;
                    if (typeof(TChar) == typeof(byte))
                    {
                        AppendEscapedUtf8Text(Unsafe.As<RecordBuffer<byte>>(output), text, quoted, encoding!);
                    }
                    else
                    {
                        AppendEscapedRecordText(Unsafe.As<RecordBuffer<char>>(output), text, ',');
                    }

                    free = output.GetSpan(InvariantFieldSpace);
                    break;
                case DataReaderFieldKind.Int32:
                    position += FormatInt32(fields.GetInt32(i), destination);
                    break;
                case DataReaderFieldKind.Int64:
                    position += FormatInt64(fields.GetInt64(i), destination);
                    break;
                case DataReaderFieldKind.Double:
                    double number = fields.GetDouble(i);
                    int formatted = FormatDouble(number, destination);
                    position += formatted > 0 ? formatted : FormatRuntimeScalar(number, destination);
                    break;
                case DataReaderFieldKind.Decimal:
                    decimal amount = fields.GetDecimal(i);
                    position += TryFormatSmallDecimal(amount, destination, out int written) ? written : FormatRuntimeScalar(amount, destination);
                    break;
                case DataReaderFieldKind.Boolean:
                    string boolean = fields.GetBoolean(i) ? "True" : "False";
                    for (var unit = 0; unit < boolean.Length; unit++)
                    {
                        destination[unit] = CsvShortestDouble.Unit<TChar>(boolean[unit]);
                    }

                    position += boolean.Length;
                    break;
                default:
                    position += FormatRuntimeScalar(fields.GetDateTime(i), destination);
                    break;
            }
        }

        output.Length += position;
        output.Write(newLine);
    }

    /// <summary>
    /// Copies text that needs no quotes after a comma when it fits: as is in UTF-16, and in UTF-8 when it is ASCII.
    /// False leaves the text to the escaping writers, with <paramref name="quoted"/> set when it is known to need quotes.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool TryCopyPlainText<TChar>(string text, Span<TChar> destination, out bool quoted)
        where TChar : unmanaged
    {
        quoted = false;
        if (text.Length > destination.Length)
        {
            return false;
        }

        if (text.Length <= ShortTextLength)
        {
            // Short text is copied and checked in one pass; the vectorized searches cost more to start than they save.
            ref TChar target = ref MemoryMarshal.GetReference(destination);
            for (var index = 0; index < text.Length; index++)
            {
                char character = text[index];
                if (character < 64)
                {
                    if ((CommaQuoteMask >> character & 1) != 0)
                    {
                        quoted = true;
                        return false;
                    }
                }
                else if (typeof(TChar) == typeof(byte) && character >= 128)
                {
                    return false;
                }

                Unsafe.Add(ref target, index) = CsvShortestDouble.Unit<TChar>(character);
            }

            return true;
        }

        if (NeedsQuotes(text.AsSpan(), ','))
        {
            quoted = true;
            return false;
        }

        if (typeof(TChar) == typeof(byte))
        {
            return Ascii.FromUtf16(text, MemoryMarshal.Cast<TChar, byte>(destination), out _) == OperationStatus.Done;
        }

        text.AsSpan().CopyTo(MemoryMarshal.Cast<TChar, char>(destination));
        return true;
    }

    private const int ShortTextLength = 16;

    /// <summary>Bit c is set for the characters below 64 that make text after a comma need quotes.</summary>
    private const ulong CommaQuoteMask = (1UL << ',') | (1UL << '"') | (1UL << '\r') | (1UL << '\n');

    /// <summary>
    /// Writes text in UTF-8 under the default quoting rule. ASCII text is narrowed between its quotes; other text goes
    /// through the destination's own encoding, so invalid UTF-16 gets the replacement or the exception it gives.
    /// </summary>
    private static void AppendEscapedUtf8Text(RecordBuffer<byte> output, string text, bool quoted, Encoding encoding)
    {
        ReadOnlySpan<char> rest = text.AsSpan();
        if (!quoted && !NeedsQuotes(rest, ','))
        {
            AppendEncoded(output, rest, encoding);
            return;
        }

        // Nothing is committed until the closing quote, so non-ASCII text can start again on the encoder path.
        Span<byte> destination = output.GetSpan(rest.Length * 2 + 2);
        int position = 0;
        destination[position++] = (byte)'"';
        ReadOnlySpan<char> remaining = rest;
        while (true)
        {
            int quote = remaining.IndexOf('"');
            ReadOnlySpan<char> segment = quote < 0 ? remaining : remaining[..(quote + 1)];
            if (Ascii.FromUtf16(segment, destination.Slice(position), out int written) != OperationStatus.Done)
            {
                AppendEncodedQuoted(output, rest, encoding);
                return;
            }

            position += written;
            if (quote < 0)
            {
                break;
            }

            destination[position++] = (byte)'"';
            remaining = remaining[(quote + 1)..];
        }

        destination[position++] = (byte)'"';
        output.Length += position;
    }

    private static void AppendEncodedQuoted(RecordBuffer<byte> output, ReadOnlySpan<char> rest, Encoding encoding)
    {
        output.Write((byte)'"');
        int quote = rest.IndexOf('"');
        while (quote >= 0)
        {
            AppendEncoded(output, rest[..(quote + 1)], encoding);
            output.Write((byte)'"');
            rest = rest[(quote + 1)..];
            quote = rest.IndexOf('"');
        }

        AppendEncoded(output, rest, encoding);
        output.Write((byte)'"');
    }

    private static void AppendEncoded(RecordBuffer<byte> output, ReadOnlySpan<char> text, Encoding encoding)
    {
        Span<byte> destination = output.GetSpan(encoding.GetMaxByteCount(text.Length));
        output.Length += encoding.GetBytes(text, destination);
    }

    /// <summary>Field reads for <see cref="AppendInvariantRecord"/>; <see cref="IsNull"/> comes first for each field.</summary>
    internal interface IRecordFields
    {
        bool IsNull(int ordinal);
        string GetString(int ordinal);
        int GetInt32(int ordinal);
        long GetInt64(int ordinal);
        double GetDouble(int ordinal);
        decimal GetDecimal(int ordinal);
        bool GetBoolean(int ordinal);
        DateTime GetDateTime(int ordinal);
    }

    /// <summary>Typed reads, after a null check.</summary>
    internal readonly struct TypedRecordFields(IDataRecord reader) : IRecordFields
    {
        public bool IsNull(int ordinal) => reader.IsDBNull(ordinal);
        public string GetString(int ordinal) => reader.GetString(ordinal);
        public int GetInt32(int ordinal) => reader.GetInt32(ordinal);
        public long GetInt64(int ordinal) => reader.GetInt64(ordinal);
        public double GetDouble(int ordinal) => reader.GetDouble(ordinal);
        public decimal GetDecimal(int ordinal) => reader.GetDecimal(ordinal);
        public bool GetBoolean(int ordinal) => reader.GetBoolean(ordinal);
        public DateTime GetDateTime(int ordinal) => reader.GetDateTime(ordinal);
    }

    /// <summary>
    /// One boxed read per field for a data table reader, which looks the row up on every call; a table column
    /// holds only its declared type or DBNull, and any other value fails the cast.
    /// </summary>
    internal struct BoxedRecordFields(IDataRecord reader) : IRecordFields
    {
        private object? _value;

        public bool IsNull(int ordinal) => (_value = reader.GetValue(ordinal)) is DBNull;
        public readonly string GetString(int ordinal) => (string)_value!;
        public readonly int GetInt32(int ordinal) => (int)_value!;
        public readonly long GetInt64(int ordinal) => (long)_value!;
        public readonly double GetDouble(int ordinal) => (double)_value!;
        public readonly decimal GetDecimal(int ordinal) => (decimal)_value!;
        public readonly bool GetBoolean(int ordinal) => (bool)_value!;
        public readonly DateTime GetDateTime(int ordinal) => (DateTime)_value!;
    }

    /// <summary>Room a field of <see cref="AppendInvariantRecord"/> may need: a delimiter and any invariant scalar.</summary>
    private const int InvariantFieldSpace = 66;

    /// <summary>The runtime's invariant text of a scalar, which never needs quotes after a comma.</summary>
    private static int FormatRuntimeScalar<T, TChar>(T value, Span<TChar> destination)
        where T : ISpanFormattable, IUtf8SpanFormattable
        where TChar : unmanaged
    {
        int written;
        bool formatted = typeof(TChar) == typeof(byte)
            ? value.TryFormat(MemoryMarshal.Cast<TChar, byte>(destination), out written, default, CultureInfo.InvariantCulture)
            : value.TryFormat(MemoryMarshal.Cast<TChar, char>(destination), out written, default, CultureInfo.InvariantCulture);
        if (!formatted)
        {
            throw new NotSupportedException("The value does not fit the record buffer.");
        }

        return written;
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
                // A null marker keeps the general path; an absent one writes nothing.
                return !formatted || nullValue is null;
            }

            switch (fieldKind)
            {
                case DataReaderFieldKind.String:
                    AppendEscapedRecordText(output, boxedValues ? (string)boxed! : reader.GetString(ordinal), delimiter);
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
                    AppendEscapedRecordText(output, (boxedValues ? (bool)boxed! : reader.GetBoolean(ordinal)) ? "True" : "False", delimiter);
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

                    Span<char> characters = stackalloc char[64];
                    if (!TryFormatDate(value, characters, out var written, format, culture))
                    {
                        return false;
                    }

                    AppendEscapedRecordText(output, characters[..written], delimiter);
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

    /// <summary>Formats a known scalar as the general path does: unescaped for invariant comma output, escaped otherwise.</summary>
    private static bool AppendRecordScalar<T>(RecordBuffer<char> output, T value, char delimiter, CultureInfo culture, bool invariantComma)
        where T : ISpanFormattable
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
            output.Write(text[..(quote + 1)]);
            output.Write('"');
            text = text[(quote + 1)..];
            quote = text.IndexOf('"');
        }

        output.Write(text);
        output.Write('"');
    }
}
#endif
