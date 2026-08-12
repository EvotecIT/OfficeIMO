#nullable enable

using System.Data;
using System.Globalization;
using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    // Aim below the large-object-heap threshold while amortizing TextWriter
    // calls. A single exceptionally large record can still exceed this size.
    internal const int DataReaderFlushThreshold = 32 * 1024;

#if NET6_0_OR_GREATER
    internal enum DataReaderFieldKind : byte
    {
        Object,
        String,
        Boolean,
        Decimal,
        Int32,
        DateTime,
        Double,
        Int64,
        DateTimeOffset,
        Guid,
        TimeSpan,
        Single,
        Byte,
        SByte,
        Int16,
        UInt16,
        UInt32,
        UInt64,
        DateOnly,
        TimeOnly
    }

    internal static DataReaderFieldKind[]? TryCreateDataReaderFieldKinds(IDataRecord reader)
    {
        var fieldKinds = new DataReaderFieldKind[reader.FieldCount];
        try
        {
            for (var i = 0; i < fieldKinds.Length; i++)
            {
                var reportedType = reader.GetFieldType(i);
                var fieldType = Nullable.GetUnderlyingType(reportedType) ?? reportedType;
                fieldKinds[i] = GetDataReaderFieldKind(fieldType);
            }
        }
        catch (NotSupportedException)
        {
            return null;
        }
        catch (NotImplementedException)
        {
            return null;
        }

        return fieldKinds;
    }

    internal static void AppendDataReaderRecordBufferedDefault(
        StringBuilder buffer,
        IDataRecord reader,
        DataReaderFieldKind[] fieldKinds,
        char delimiter,
        string newLine,
        CultureInfo culture)
    {
        for (var i = 0; i < fieldKinds.Length; i++)
        {
            if (i > 0)
            {
                buffer.Append(delimiter);
            }

            var fieldKind = fieldKinds[i];
            if (fieldKind == DataReaderFieldKind.Object || !TryAppendTypedDataReaderValue(
                    buffer,
                    reader,
                    i,
                    fieldKind,
                    delimiter,
                    culture))
            {
                fieldKinds[i] = DataReaderFieldKind.Object;
                AppendDataReaderObjectValue(buffer, reader.GetValue(i), delimiter, culture);
            }
        }

        buffer.Append(newLine);
    }

    private static bool TryAppendTypedDataReaderValue(
        StringBuilder buffer,
        IDataRecord reader,
        int ordinal,
        DataReaderFieldKind fieldKind,
        char delimiter,
        CultureInfo culture)
    {
        try
        {
            if (reader.IsDBNull(ordinal))
            {
                return true;
            }

            switch (fieldKind)
            {
                case DataReaderFieldKind.String:
                    WriteEscapedDefault(buffer, reader.GetString(ordinal), delimiter);
                    return true;
                case DataReaderFieldKind.Boolean:
                    var boolean = reader.GetBoolean(ordinal);
                    if (delimiter == ',')
                    {
                        buffer.Append(boolean ? "True" : "False");
                    }
                    else
                    {
                        WriteEscapedDefault(buffer, boolean ? "True" : "False", delimiter);
                    }

                    return true;
                case DataReaderFieldKind.Decimal:
                    AppendKnownValueDefault(buffer, reader.GetDecimal(ordinal), delimiter, culture);
                    return true;
                case DataReaderFieldKind.Int32:
                    AppendKnownValueDefault(buffer, reader.GetInt32(ordinal), delimiter, culture);
                    return true;
                case DataReaderFieldKind.DateTime:
                    AppendKnownValueDefault(buffer, reader.GetDateTime(ordinal), delimiter, culture);
                    return true;
                case DataReaderFieldKind.Double:
                    AppendKnownValueDefault(buffer, reader.GetDouble(ordinal), delimiter, culture);
                    return true;
                case DataReaderFieldKind.Int64:
                    AppendKnownValueDefault(buffer, reader.GetInt64(ordinal), delimiter, culture);
                    return true;
                case DataReaderFieldKind.Guid:
                    AppendKnownValueDefault(buffer, reader.GetGuid(ordinal), delimiter, culture);
                    return true;
                case DataReaderFieldKind.Single:
                    AppendKnownValueDefault(buffer, reader.GetFloat(ordinal), delimiter, culture);
                    return true;
                case DataReaderFieldKind.Byte:
                    AppendKnownValueDefault(buffer, reader.GetByte(ordinal), delimiter, culture);
                    return true;
                case DataReaderFieldKind.Int16:
                    AppendKnownValueDefault(buffer, reader.GetInt16(ordinal), delimiter, culture);
                    return true;
                default:
                    return false;
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

    private static void AppendDataReaderObjectValue(
        StringBuilder buffer,
        object? value,
        char delimiter,
        CultureInfo culture)
    {
        if (value is null || ReferenceEquals(value, DBNull.Value))
        {
            return;
        }

        AppendEscapedValueDefault(buffer, value, delimiter, culture);
    }

    internal static void FlushBufferedContent(TextWriter writer, StringBuilder buffer)
    {
        if (buffer.Length == 0)
        {
            return;
        }

        writer.Write(buffer);
        buffer.Clear();
    }

    private static DataReaderFieldKind GetDataReaderFieldKind(Type fieldType)
    {
        if (fieldType == typeof(string)) return DataReaderFieldKind.String;
        if (fieldType == typeof(bool)) return DataReaderFieldKind.Boolean;
        if (fieldType == typeof(decimal)) return DataReaderFieldKind.Decimal;
        if (fieldType == typeof(int)) return DataReaderFieldKind.Int32;
        if (fieldType == typeof(DateTime)) return DataReaderFieldKind.DateTime;
        if (fieldType == typeof(double)) return DataReaderFieldKind.Double;
        if (fieldType == typeof(long)) return DataReaderFieldKind.Int64;
        if (fieldType == typeof(DateTimeOffset)) return DataReaderFieldKind.DateTimeOffset;
        if (fieldType == typeof(Guid)) return DataReaderFieldKind.Guid;
        if (fieldType == typeof(TimeSpan)) return DataReaderFieldKind.TimeSpan;
        if (fieldType == typeof(float)) return DataReaderFieldKind.Single;
        if (fieldType == typeof(byte)) return DataReaderFieldKind.Byte;
        if (fieldType == typeof(sbyte)) return DataReaderFieldKind.SByte;
        if (fieldType == typeof(short)) return DataReaderFieldKind.Int16;
        if (fieldType == typeof(ushort)) return DataReaderFieldKind.UInt16;
        if (fieldType == typeof(uint)) return DataReaderFieldKind.UInt32;
        if (fieldType == typeof(ulong)) return DataReaderFieldKind.UInt64;
        if (fieldType == typeof(DateOnly)) return DataReaderFieldKind.DateOnly;
        if (fieldType == typeof(TimeOnly)) return DataReaderFieldKind.TimeOnly;
        return DataReaderFieldKind.Object;
    }

    private static void AppendKnownValueDefault<T>(
        StringBuilder buffer,
        T value,
        char delimiter,
        CultureInfo culture)
        where T : ISpanFormattable
    {
        if (delimiter == ',' && ReferenceEquals(culture, CultureInfo.InvariantCulture))
        {
            buffer.Append(CultureInfo.InvariantCulture, $"{value}");
            return;
        }

        Span<char> destination = stackalloc char[128];
        if (!value.TryFormat(destination, out var charsWritten, default, culture))
        {
            WriteEscapedDefault(buffer, FormatValue(value, culture), delimiter);
            return;
        }

        var formatted = destination[..charsWritten];
        AppendEscapedSpanDefault(buffer, formatted, delimiter);
    }
#endif
}
