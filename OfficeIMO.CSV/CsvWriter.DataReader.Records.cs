#if NET8_0_OR_GREATER
using System.Buffers;
using System.Data;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
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

    /// <summary>
    /// Column layouts of <see cref="AppendInvariantRecord{TFields, TChar, TLayout}"/>. Each layout is its own
    /// value-type instantiation, so text-only, scalar-only and mixed rows get separate machine code and separate
    /// tiering profiles: writing one shape first does not decide how the others are compiled.
    /// </summary>
    internal interface IInvariantRecordLayout { }

    /// <summary>Every field is a string.</summary>
    internal struct TextRecordLayout : IInvariantRecordLayout { }

    /// <summary>No field is a string.</summary>
    internal struct ScalarRecordLayout : IInvariantRecordLayout { }

    /// <summary>Strings and scalars together.</summary>
    internal struct MixedRecordLayout : IInvariantRecordLayout { }

    internal enum InvariantRecordLayout : byte
    {
        None,
        Text,
        Scalar,
        Mixed
    }

    /// <summary>
    /// The layout of a record <see cref="AppendInvariantRecord{TFields, TChar, TLayout}"/> writes, or None when a
    /// field kind is not one it writes.
    /// </summary>
    internal static InvariantRecordLayout GetInvariantRecordLayout(DataReaderFieldKind[] fieldKinds)
    {
        bool text = false;
        bool scalar = false;
        foreach (DataReaderFieldKind kind in fieldKinds)
        {
            if (kind == DataReaderFieldKind.String)
            {
                text = true;
            }
            else if (kind is DataReaderFieldKind.Int32 or DataReaderFieldKind.Int64 or DataReaderFieldKind.Double
                or DataReaderFieldKind.Decimal or DataReaderFieldKind.Boolean or DataReaderFieldKind.DateTime)
            {
                scalar = true;
            }
            else
            {
                return InvariantRecordLayout.None;
            }
        }

        return text ? scalar ? InvariantRecordLayout.Mixed : InvariantRecordLayout.Text : InvariantRecordLayout.Scalar;
    }

    /// <summary>
    /// Appends a record of strings, integers, doubles, decimals, booleans and dates for comma output in the
    /// invariant culture without value formatting options, as <see cref="AppendDataReaderRecord"/> writes it,
    /// in one pass with no per-field exception handling. A reader that rejects a typed read throws out of here;
    /// the caller then rewinds the row and writes it through <see cref="AppendDataReaderRecord"/>.
    /// <typeparamref name="TLayout"/> matches <see cref="GetInvariantRecordLayout"/> of the field kinds.
    /// </summary>
    internal static void AppendInvariantRecord<TFields, TChar, TLayout>(
        RecordBuffer<TChar> output, ref TFields fields, DataReaderFieldKind[] fieldKinds, ReadOnlySpan<TChar> newLine, Encoding? encoding)
        where TFields : struct, IRecordFields
        where TChar : unmanaged
        where TLayout : struct, IInvariantRecordLayout
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

            if (typeof(TLayout) == typeof(TextRecordLayout))
            {
                AppendInvariantText(output, ref free, ref position, fields.GetString(i), encoding, quotesExpected: true);
                continue;
            }

            Span<TChar> destination = free.Slice(position);
            switch (fieldKinds[i])
            {
                // Rows without strings leave this case out of their code.
                case DataReaderFieldKind.String when typeof(TLayout) != typeof(ScalarRecordLayout):
                    AppendInvariantText(output, ref free, ref position, fields.GetString(i), encoding, quotesExpected: false);
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
    /// Writes one text field of <see cref="AppendInvariantRecord{TFields, TChar, TLayout}"/> at
    /// <paramref name="position"/> in <paramref name="free"/>, committing and renewing the span when it needs more
    /// room. Short text is escaped in one pass; <paramref name="quotesExpected"/> skips the plain copy that suits
    /// text that seldom needs quotes.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static void AppendInvariantText<TChar>(
        RecordBuffer<TChar> output, ref Span<TChar> free, ref int position, string text, Encoding? encoding, bool quotesExpected)
        where TChar : unmanaged
    {
        if (text.Length <= EscapedShortTextLength)
        {
            int room = text.Length * 2 + 2;
            if (free.Length - position < room)
            {
                output.Length += position;
                position = 0;
                free = output.GetSpan(Math.Max(room, InvariantFieldSpace));
            }

            Span<TChar> destination = free.Slice(position);
            if (!quotesExpected && text.Length <= ShortTextLength && TryCopyPlainText(text, destination, out _))
            {
                position += text.Length;
                return;
            }

            if (TryWriteEscapedShortText(text, destination, out int written))
            {
                position += written;
                return;
            }
        }
        else if (TryCopyPlainText(text, free.Slice(position), out _))
        {
            position += text.Length;
            return;
        }

        output.Length += position;
        position = 0;
        if (typeof(TChar) == typeof(byte))
        {
            AppendEscapedUtf8Text(Unsafe.As<RecordBuffer<byte>>(output), text, quoted: false, encoding!);
        }
        else
        {
            AppendEscapedRecordText(Unsafe.As<RecordBuffer<char>>(output), text, ',');
        }

        free = output.GetSpan(InvariantFieldSpace);
    }

    /// <summary>Text up to this length is escaped in one pass into room for twice its length and two quotes.</summary>
    internal const int EscapedShortTextLength = 64;

    /// <summary>
    /// Escapes short text after a comma in one pass: copies it one unit along with its quotes doubled, then adds
    /// the enclosing quotes when it holds a comma, quote or line break, or moves it back when it does not.
    /// <paramref name="destination"/> holds at least twice the text length and two more units. False for
    /// non-ASCII text in UTF-8, which the caller encodes through its encoding.
    /// </summary>
    internal static bool TryWriteEscapedShortText<TChar>(string text, Span<TChar> destination, out int written)
        where TChar : unmanaged
    {
        if (Vector128.IsHardwareAccelerated && text.Length >= Vector128<ushort>.Count)
        {
            // Text without quote characters, the usual case, is checked and copied eight units at a time.
            if (TryWriteEscapedUnquotedText(text, destination, out written, out bool hasQuote))
            {
                return true;
            }

            if (!hasQuote)
            {
                // Non-ASCII text in UTF-8.
                return false;
            }
        }

        ref TChar target = ref MemoryMarshal.GetReference(destination);
        int position = 1;
        bool special = false;
        for (var index = 0; index < text.Length; index++)
        {
            char character = text[index];
            if (character < 64)
            {
                if ((CommaQuoteMask >> character & 1) != 0)
                {
                    special = true;
                    if (character == '"')
                    {
                        Unsafe.Add(ref target, position++) = CsvShortestDouble.Unit<TChar>('"');
                    }
                }
            }
            else if (typeof(TChar) == typeof(byte) && character >= 128)
            {
                written = 0;
                return false;
            }

            Unsafe.Add(ref target, position++) = CsvShortestDouble.Unit<TChar>(character);
        }

        if (special)
        {
            target = CsvShortestDouble.Unit<TChar>('"');
            Unsafe.Add(ref target, position++) = CsvShortestDouble.Unit<TChar>('"');
            written = position;
        }
        else
        {
            destination.Slice(1, position - 1).CopyTo(destination);
            written = position - 1;
        }

        return true;
    }

    /// <summary>
    /// The vector path of <see cref="TryWriteEscapedShortText{TChar}"/> for text of at least eight units: finds commas,
    /// quotes, line breaks and, in UTF-8, non-ASCII characters eight at a time (the last eight overlapping the ones
    /// before), then copies the text after an opening quote when it needs one. False with <paramref name="hasQuote"/>
    /// set leaves text holding a quote character to the one-pass loop, which doubles it; false without it is
    /// non-ASCII text in UTF-8, which the caller encodes.
    /// </summary>
    private static bool TryWriteEscapedUnquotedText<TChar>(string text, Span<TChar> destination, out int written, out bool hasQuote)
        where TChar : unmanaged
    {
        written = 0;
        hasQuote = false;
        int length = text.Length;
        if (destination.Length < length + 2)
        {
            // Left to the one-pass loop, which relies on the caller's room for the escaped text.
            hasQuote = true;
            return false;
        }

        ref ushort source = ref Unsafe.As<char, ushort>(ref MemoryMarshal.GetReference(text.AsSpan()));
        nuint last = (nuint)(length - Vector128<ushort>.Count);
        Vector128<ushort> quotes = Vector128<ushort>.Zero;
        Vector128<ushort> specials = Vector128<ushort>.Zero;
        Vector128<ushort> nonAscii = Vector128<ushort>.Zero;
        for (nuint offset = 0; ; offset = Math.Min(offset + (nuint)Vector128<ushort>.Count, last))
        {
            Vector128<ushort> units = Vector128.LoadUnsafe(ref source, offset);
            quotes |= Vector128.Equals(units, Vector128.Create((ushort)'"'));
            specials |= Vector128.Equals(units, Vector128.Create((ushort)',')) | Vector128.Equals(units, Vector128.Create((ushort)'\n'))
                | Vector128.Equals(units, Vector128.Create((ushort)'\r'));
            if (typeof(TChar) == typeof(byte))
            {
                nonAscii |= Vector128.GreaterThan(units, Vector128.Create((ushort)127));
            }

            if (offset == last)
            {
                break;
            }
        }

        if (typeof(TChar) == typeof(byte) && nonAscii != Vector128<ushort>.Zero)
        {
            return false;
        }

        if (quotes != Vector128<ushort>.Zero)
        {
            hasQuote = true;
            return false;
        }

        int start = specials == Vector128<ushort>.Zero ? 0 : 1;
        ref TChar target = ref Unsafe.Add(ref MemoryMarshal.GetReference(destination), start);
        for (nuint offset = 0; ; offset = Math.Min(offset + (nuint)Vector128<ushort>.Count, last))
        {
            Vector128<ushort> units = Vector128.LoadUnsafe(ref source, offset);
            if (typeof(TChar) == typeof(byte))
            {
                Unsafe.WriteUnaligned(ref Unsafe.As<TChar, byte>(ref Unsafe.Add(ref target, offset)),
                    Vector128.Narrow(units, units).AsUInt64().ToScalar());
            }
            else
            {
                units.StoreUnsafe(ref Unsafe.As<TChar, ushort>(ref target), offset);
            }

            if (offset == last)
            {
                break;
            }
        }

        if (start == 0)
        {
            written = length;
            return true;
        }

        destination[0] = CsvShortestDouble.Unit<TChar>('"');
        destination[length + 1] = CsvShortestDouble.Unit<TChar>('"');
        written = length + 2;
        return true;
    }

    /// <summary>
    /// Copies text that needs no quotes after a comma when it fits: as is in UTF-16, and in UTF-8 when it is ASCII.
    /// False leaves the text to the escaping writers, with <paramref name="quoted"/> set when it is known to need quotes.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static bool TryCopyPlainText<TChar>(string text, Span<TChar> destination, out bool quoted)
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

    internal const int ShortTextLength = 16;

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

    /// <summary>Field reads for <see cref="AppendInvariantRecord{TFields, TChar, TLayout}"/>; <see cref="IsNull"/> comes first for each field.</summary>
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

    /// <summary>Room a field of <see cref="AppendInvariantRecord{TFields, TChar, TLayout}"/> may need: a delimiter and any invariant scalar.</summary>
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
