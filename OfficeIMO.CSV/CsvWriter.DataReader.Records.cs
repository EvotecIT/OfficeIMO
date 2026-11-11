#if NET8_0_OR_GREATER
using System.Buffers;
using System.Data;
using System.Globalization;
using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    /// <summary>Growable pooled characters for completed data-reader rows.</summary>
    internal sealed class RecordBuffer : IDisposable
    {
        private char[] _characters = ArrayPool<char>.Shared.Rent(DataReaderFlushThreshold + 4096);

        internal int Length { get; set; }

        internal ReadOnlySpan<char> Written => _characters.AsSpan(0, Length);

        internal Span<char> GetSpan(int sizeHint)
        {
            if (_characters.Length - Length < sizeHint)
            {
                char[] larger = ArrayPool<char>.Shared.Rent(Math.Max(_characters.Length * 2, Length + sizeHint));
                _characters.AsSpan(0, Length).CopyTo(larger);
                ArrayPool<char>.Shared.Return(_characters);
                _characters = larger;
            }

            return _characters.AsSpan(Length);
        }

        internal void Write(char value)
        {
            GetSpan(1)[0] = value;
            Length++;
        }

        internal void Write(ReadOnlySpan<char> value)
        {
            value.CopyTo(GetSpan(value.Length));
            Length += value.Length;
        }


        public void Dispose()
        {
            ArrayPool<char>.Shared.Return(_characters, clearArray: true);
            _characters = Array.Empty<char>();
        }
    }

    /// <summary>
    /// Appends one typed data-reader record. Strings, integers, doubles, decimals, booleans and dates are
    /// formatted straight into the buffer; every other value goes through the row appenders of the general
    /// typed path, so the text is the same as that path writes.
    /// </summary>
    internal static void AppendDataReaderRecord(
        RecordBuffer output,
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
    internal static void AppendInvariantRecord<TFields>(RecordBuffer output, ref TFields fields, DataReaderFieldKind[] fieldKinds, string newLine)
        where TFields : struct, IRecordFields
    {
        for (var i = 0; i < fieldKinds.Length; i++)
        {
            if (i > 0)
            {
                output.Write(',');
            }

            if (fields.IsNull(i))
            {
                continue;
            }

            switch (fieldKinds[i])
            {
                case DataReaderFieldKind.String:
                    AppendEscapedRecordText(output, fields.GetString(i), ',');
                    break;
                case DataReaderFieldKind.Int32:
                    AppendInvariantScalar(output, fields.GetInt32(i));
                    break;
                case DataReaderFieldKind.Int64:
                    AppendInvariantScalar(output, fields.GetInt64(i));
                    break;
                case DataReaderFieldKind.Double:
                    AppendInvariantScalar(output, fields.GetDouble(i));
                    break;
                case DataReaderFieldKind.Decimal:
                    AppendInvariantScalar(output, fields.GetDecimal(i));
                    break;
                case DataReaderFieldKind.Boolean:
                    output.Write(fields.GetBoolean(i) ? "True" : "False");
                    break;
                default:
                    AppendInvariantScalar(output, fields.GetDateTime(i));
                    break;
            }
        }

        output.Write(newLine);
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

    /// <summary>Invariant text of these scalars never needs quotes after a comma.</summary>
    private static void AppendInvariantScalar<T>(RecordBuffer output, T value) where T : ISpanFormattable
    {
        if (!TryFormatInvariant(value, output.GetSpan(64), out int written))
        {
            throw new NotSupportedException("The value does not fit the record buffer.");
        }

        output.Length += written;
    }

    /// <summary>Writes the common kinds directly; false leaves the value to the general appenders.</summary>
    private static bool TryAppendRecordValue(
        RecordBuffer output,
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
    private static bool AppendRecordScalar<T>(RecordBuffer output, T value, char delimiter, CultureInfo culture, bool invariantComma)
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
    private static void AppendEscapedRecordText(RecordBuffer output, ReadOnlySpan<char> text, char delimiter)
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
