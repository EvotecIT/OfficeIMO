#nullable enable

using System;
using System.Threading;

namespace OfficeIMO.CSV;

internal interface ICsvDataReaderTextRowSource : IDisposable
{
    bool Read();

    bool Read(CancellationToken cancellationToken);

#if NET8_0_OR_GREATER
    ReadOnlySpan<char> GetSpan(int ordinal);
#endif

    string GetString(int ordinal);

    bool IsMissing(int ordinal);

    bool IsNull(int ordinal, string? nullValue);

    int CopyStringValues(object[] values, int count, string? nullValue);
}

internal interface ICsvDataReaderHeaderRowSource : ICsvDataReaderTextRowSource
{
    int FieldCount { get; }

    void SetSourceColumnCount(int sourceColumnCount);
}

#if NET8_0_OR_GREATER
internal interface ICsvAsyncDataReaderRowSource : ICsvDataReaderTextRowSource
{
    System.Threading.Tasks.ValueTask<bool> ReadAsync(CancellationToken cancellationToken);
    bool HasStaticValues { get; }
    object? GetRawValue(int ordinal);
}
#endif

internal interface ICsvDataReaderParallelBatchInfo
{
    int RowCount { get; }
}
