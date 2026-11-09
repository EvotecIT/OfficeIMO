#if NET8_0_OR_GREATER

namespace OfficeIMO.CSV;

/// <summary>Exposes an owned bounded record batch through the ordinary text reader contract.</summary>
internal sealed class CsvRecordBatchRowSource : ICsvDataReaderTextRowSource, ICsvDataReaderParallelBatchInfo, ICsvDataReaderPositionSource
{
    private readonly CsvRecordTextBatch _batch;
    private bool _prepared;
    internal CsvRecordBatchRowSource(CsvRecordTextBatch batch) => _batch = batch;
    public int RowCount => _batch.Count;
    public int? CurrentPhysicalLineNumber => _batch.CurrentPhysicalLineNumber;
    public int? CurrentPhysicalEndLineNumber => _batch.CurrentPhysicalEndLineNumber;
    public bool Read()
    {
        if (!_prepared) { _batch.PrepareDecodedText(); _prepared = true; }
        return _batch.Read();
    }
    public bool Read(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Read();
    }
    public ReadOnlySpan<char> GetSpan(int ordinal) => _batch.GetSpan(ordinal);
    public string GetString(int ordinal) => _batch.GetString(ordinal);
    public bool IsMissing(int ordinal) => _batch.IsMissing(ordinal);
    public bool IsNull(int ordinal, string? nullValue) => _batch.IsConfiguredNull(ordinal);
    public int CopyStringValues(object[] values, int count, string? nullValue)
    {
        int copied = Math.Min(count, _batch.FieldCount);
        for (int i = 0; i < copied; i++) values[i] = IsNull(i, nullValue) ? DBNull.Value : GetString(i);
        return copied;
    }
    public void Dispose() => _batch.Dispose();
}
#endif
