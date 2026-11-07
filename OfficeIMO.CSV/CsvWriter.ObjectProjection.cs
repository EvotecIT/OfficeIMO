#if NET8_0_OR_GREATER
using System.Reflection;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    /// <summary>Roots the concrete scalar formatter selected by the existing field classifier.</summary>
    internal static MethodInfo? GetPropertyScalarStageMethod(Type type) => GetDataReaderFieldKind(type) switch
    {
        DataReaderFieldKind.Decimal => CsvRowWriter.GetPropertyScalarStageMethod<decimal>(),
        DataReaderFieldKind.Int32 => CsvRowWriter.GetPropertyScalarStageMethod<int>(),
        DataReaderFieldKind.DateTime => CsvRowWriter.GetPropertyScalarStageMethod<DateTime>(),
        DataReaderFieldKind.Double => CsvRowWriter.GetPropertyScalarStageMethod<double>(),
        DataReaderFieldKind.Int64 => CsvRowWriter.GetPropertyScalarStageMethod<long>(),
        DataReaderFieldKind.DateTimeOffset => CsvRowWriter.GetPropertyScalarStageMethod<DateTimeOffset>(),
        DataReaderFieldKind.Guid => CsvRowWriter.GetPropertyScalarStageMethod<Guid>(),
        DataReaderFieldKind.TimeSpan => CsvRowWriter.GetPropertyScalarStageMethod<TimeSpan>(),
        DataReaderFieldKind.Single => CsvRowWriter.GetPropertyScalarStageMethod<float>(),
        DataReaderFieldKind.Byte => CsvRowWriter.GetPropertyScalarStageMethod<byte>(),
        DataReaderFieldKind.SByte => CsvRowWriter.GetPropertyScalarStageMethod<sbyte>(),
        DataReaderFieldKind.Int16 => CsvRowWriter.GetPropertyScalarStageMethod<short>(),
        DataReaderFieldKind.UInt16 => CsvRowWriter.GetPropertyScalarStageMethod<ushort>(),
        DataReaderFieldKind.UInt32 => CsvRowWriter.GetPropertyScalarStageMethod<uint>(),
        DataReaderFieldKind.UInt64 => CsvRowWriter.GetPropertyScalarStageMethod<ulong>(),
        DataReaderFieldKind.DateOnly => CsvRowWriter.GetPropertyScalarStageMethod<DateOnly>(),
        DataReaderFieldKind.TimeOnly => CsvRowWriter.GetPropertyScalarStageMethod<TimeOnly>(),
        _ => null,
    };
}
#endif
