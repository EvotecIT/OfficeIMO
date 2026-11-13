#if NET8_0_OR_GREATER
using System.Reflection;

namespace OfficeIMO.CSV;

internal static partial class CsvWriter
{
    /// <summary>Roots the concrete scalar formatter selected by the existing field classifier.</summary>
    internal static MethodInfo? GetPropertyScalarStageMethod(Type type, CsvRowWriter.ProjectedStaging staging) => GetDataReaderFieldKind(type) switch
    {
        DataReaderFieldKind.Decimal => CsvRowWriter.GetPropertyScalarStageMethod<decimal>(staging),
        DataReaderFieldKind.Int32 => CsvRowWriter.GetPropertyScalarStageMethod<int>(staging),
        DataReaderFieldKind.DateTime => CsvRowWriter.GetPropertyScalarStageMethod<DateTime>(staging),
        DataReaderFieldKind.Double => CsvRowWriter.GetPropertyScalarStageMethod<double>(staging),
        DataReaderFieldKind.Int64 => CsvRowWriter.GetPropertyScalarStageMethod<long>(staging),
        DataReaderFieldKind.DateTimeOffset => CsvRowWriter.GetPropertyScalarStageMethod<DateTimeOffset>(staging),
        DataReaderFieldKind.Guid => CsvRowWriter.GetPropertyScalarStageMethod<Guid>(staging),
        DataReaderFieldKind.TimeSpan => CsvRowWriter.GetPropertyScalarStageMethod<TimeSpan>(staging),
        DataReaderFieldKind.Single => CsvRowWriter.GetPropertyScalarStageMethod<float>(staging),
        DataReaderFieldKind.Byte => CsvRowWriter.GetPropertyScalarStageMethod<byte>(staging),
        DataReaderFieldKind.SByte => CsvRowWriter.GetPropertyScalarStageMethod<sbyte>(staging),
        DataReaderFieldKind.Int16 => CsvRowWriter.GetPropertyScalarStageMethod<short>(staging),
        DataReaderFieldKind.UInt16 => CsvRowWriter.GetPropertyScalarStageMethod<ushort>(staging),
        DataReaderFieldKind.UInt32 => CsvRowWriter.GetPropertyScalarStageMethod<uint>(staging),
        DataReaderFieldKind.UInt64 => CsvRowWriter.GetPropertyScalarStageMethod<ulong>(staging),
        DataReaderFieldKind.DateOnly => CsvRowWriter.GetPropertyScalarStageMethod<DateOnly>(staging),
        DataReaderFieldKind.TimeOnly => CsvRowWriter.GetPropertyScalarStageMethod<TimeOnly>(staging),
        _ => null,
    };
}
#endif
