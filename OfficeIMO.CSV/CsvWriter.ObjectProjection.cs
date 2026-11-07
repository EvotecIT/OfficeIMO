#if NET8_0_OR_GREATER
using System.Reflection;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    /// <summary>Roots the concrete scalar formatter selected by the existing field classifier.</summary>
    internal static MethodInfo? GetPropertyScalarStageMethod(Type type, bool useSpanBuffer = false) => GetDataReaderFieldKind(type) switch
    {
        DataReaderFieldKind.Decimal => CsvRowWriter.GetPropertyScalarStageMethod<decimal>(useSpanBuffer),
        DataReaderFieldKind.Int32 => CsvRowWriter.GetPropertyScalarStageMethod<int>(useSpanBuffer),
        DataReaderFieldKind.DateTime => CsvRowWriter.GetPropertyScalarStageMethod<DateTime>(useSpanBuffer),
        DataReaderFieldKind.Double => CsvRowWriter.GetPropertyScalarStageMethod<double>(useSpanBuffer),
        DataReaderFieldKind.Int64 => CsvRowWriter.GetPropertyScalarStageMethod<long>(useSpanBuffer),
        DataReaderFieldKind.DateTimeOffset => CsvRowWriter.GetPropertyScalarStageMethod<DateTimeOffset>(useSpanBuffer),
        DataReaderFieldKind.Guid => CsvRowWriter.GetPropertyScalarStageMethod<Guid>(useSpanBuffer),
        DataReaderFieldKind.TimeSpan => CsvRowWriter.GetPropertyScalarStageMethod<TimeSpan>(useSpanBuffer),
        DataReaderFieldKind.Single => CsvRowWriter.GetPropertyScalarStageMethod<float>(useSpanBuffer),
        DataReaderFieldKind.Byte => CsvRowWriter.GetPropertyScalarStageMethod<byte>(useSpanBuffer),
        DataReaderFieldKind.SByte => CsvRowWriter.GetPropertyScalarStageMethod<sbyte>(useSpanBuffer),
        DataReaderFieldKind.Int16 => CsvRowWriter.GetPropertyScalarStageMethod<short>(useSpanBuffer),
        DataReaderFieldKind.UInt16 => CsvRowWriter.GetPropertyScalarStageMethod<ushort>(useSpanBuffer),
        DataReaderFieldKind.UInt32 => CsvRowWriter.GetPropertyScalarStageMethod<uint>(useSpanBuffer),
        DataReaderFieldKind.UInt64 => CsvRowWriter.GetPropertyScalarStageMethod<ulong>(useSpanBuffer),
        DataReaderFieldKind.DateOnly => CsvRowWriter.GetPropertyScalarStageMethod<DateOnly>(useSpanBuffer),
        DataReaderFieldKind.TimeOnly => CsvRowWriter.GetPropertyScalarStageMethod<TimeOnly>(useSpanBuffer),
        _ => null,
    };
}
#endif
