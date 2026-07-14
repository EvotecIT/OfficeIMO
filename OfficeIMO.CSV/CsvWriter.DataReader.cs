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

            var value = reader.GetValue(i);
            if (value is null || ReferenceEquals(value, DBNull.Value))
            {
                continue;
            }

            switch (fieldKinds[i])
            {
                case DataReaderFieldKind.String when value is string text:
                    WriteEscapedDefault(buffer, text, delimiter);
                    break;
                case DataReaderFieldKind.Boolean when value is bool boolean:
                    if (delimiter == ',')
                    {
                        buffer.Append(boolean ? "True" : "False");
                    }
                    else
                    {
                        WriteEscapedDefault(buffer, boolean ? "True" : "False", delimiter);
                    }
                    break;
                case DataReaderFieldKind.Decimal when value is decimal number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.Int32 when value is int number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.DateTime when value is DateTime dateTime:
                    AppendKnownValueDefault(buffer, dateTime, delimiter, culture);
                    break;
                case DataReaderFieldKind.Double when value is double number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.Int64 when value is long number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.DateTimeOffset when value is DateTimeOffset dateTimeOffset:
                    AppendKnownValueDefault(buffer, dateTimeOffset, delimiter, culture);
                    break;
                case DataReaderFieldKind.Guid when value is Guid guid:
                    AppendKnownValueDefault(buffer, guid, delimiter, culture);
                    break;
                case DataReaderFieldKind.TimeSpan when value is TimeSpan timeSpan:
                    AppendKnownValueDefault(buffer, timeSpan, delimiter, culture);
                    break;
                case DataReaderFieldKind.Single when value is float number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.Byte when value is byte number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.SByte when value is sbyte number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.Int16 when value is short number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.UInt16 when value is ushort number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.UInt32 when value is uint number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.UInt64 when value is ulong number:
                    AppendKnownValueDefault(buffer, number, delimiter, culture);
                    break;
                case DataReaderFieldKind.DateOnly when value is DateOnly dateOnly:
                    AppendKnownValueDefault(buffer, dateOnly, delimiter, culture);
                    break;
                case DataReaderFieldKind.TimeOnly when value is TimeOnly timeOnly:
                    AppendKnownValueDefault(buffer, timeOnly, delimiter, culture);
                    break;
                default:
                    AppendEscapedValueDefault(buffer, value, delimiter, culture);
                    break;
            }
        }

        buffer.Append(newLine);
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
            AppendEscapedValueDefault(buffer, value, delimiter, culture);
            return;
        }

        var formatted = destination[..charsWritten];
        AppendEscapedSpanDefault(buffer, formatted, delimiter);
    }
#endif
}
