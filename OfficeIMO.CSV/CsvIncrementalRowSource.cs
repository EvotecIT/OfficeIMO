#if NET8_0_OR_GREATER
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

internal sealed class CsvIncrementalRowSource : ICsvAsyncDataReaderRowSource, ICsvDataReaderPositionSource
{
    internal readonly record struct BufferedRecord(IReadOnlyList<string> Values, int Start, int End);
    private readonly CsvParser.IncrementalRecords _records;
    private readonly CsvParsingSettings _options;
    private readonly CancellationTokenSource _lifetime;
    private readonly CancellationToken _openingToken, _loadToken;
    private readonly int _headerCount;
    private readonly Queue<BufferedRecord> _buffered;
    private IReadOnlyList<string> _current = Array.Empty<string>();
    private bool _disposed;
    private bool _failed;
    private int _sourceValueCount;
    private readonly object?[] _staticValues;

    internal CsvIncrementalRowSource(CsvParser.IncrementalRecords records, CsvParsingSettings options,
        CancellationTokenSource lifetime, CancellationToken openingToken, CancellationToken loadToken,
        int headerCount, Queue<BufferedRecord> buffered)
    {
        _records = records;
        _options = options;
        _lifetime = lifetime;
        _openingToken = openingToken;
        _loadToken = loadToken;
        _headerCount = headerCount;
        _buffered = buffered;
        _staticValues = options.StaticColumns?.Values.ToArray() ?? Array.Empty<object?>();
    }

    public int? CurrentPhysicalLineNumber { get; private set; }
    public int? CurrentPhysicalEndLineNumber { get; private set; }
    public bool Read() => Read(_options.CancellationToken);
    public bool Read(CancellationToken token) => AdvanceAsync(false, token).GetAwaiter().GetResult();
    public ValueTask<bool> ReadAsync(CancellationToken token) => AdvanceAsync(true, token);

    private ValueTask<bool> AdvanceAsync(bool asynchronous, CancellationToken token)
    {
        CancellationTokenSource? operation = null;
        try
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_failed) throw new InvalidOperationException("The CSV reader cannot continue after a failed advance.");
            try
            {
                token.ThrowIfCancellationRequested();
                _options.CancellationToken.ThrowIfCancellationRequested();
                // A pending read keeps its linked token alive until completion.
                operation = LinkReadCancellation(token);
                var effective = operation?.Token ?? _options.CancellationToken;
                if (_buffered.Count > 0)
                {
                    SetCurrent(_buffered.Dequeue());
                    return new ValueTask<bool>(true);
                }
                ValueTask<bool> read = _records.ReadAsync(asynchronous, effective, reuseValues: true);
                if (read.IsCompletedSuccessfully)
                    return new ValueTask<bool>(CompleteRecordRead(read.Result, effective));

                var pendingOperation = operation;
                operation = null;
                return AwaitRecordReadAsync(read, effective, pendingOperation);
            }
            catch { _failed = true; throw; }
        }
        catch (Exception exception) { return CsvParser.PropagateReadFailureAsync(exception); }
        finally { operation?.Dispose(); }
    }

    private async ValueTask<bool> AwaitRecordReadAsync(ValueTask<bool> read,
        CancellationToken effective, CancellationTokenSource? operation)
    {
        try { return CompleteRecordRead(await read.ConfigureAwait(false), effective); }
        catch { _failed = true; throw; }
        finally { operation?.Dispose(); }
    }

    private bool CompleteRecordRead(bool available, CancellationToken effective)
    {
        if (!available)
        {
            _current = Array.Empty<string>();
            CurrentPhysicalLineNumber = CurrentPhysicalEndLineNumber = null;
            return false;
        }
        effective.ThrowIfCancellationRequested();
        SetCurrent(new BufferedRecord(_records.Current.Values, _records.StartLine, _records.EndLine));
        return true;
    }

    internal bool CanCaptureRawRecords => _buffered.Count == 0 && !HasStaticValues &&
        _headerCount > 0 && _records.CanCaptureRawRecords;

    [System.Runtime.CompilerServices.AsyncMethodBuilder(typeof(System.Runtime.CompilerServices.PoolingAsyncValueTaskMethodBuilder<>))]
    internal async ValueTask<(int Count, bool ReachedEnd)> CaptureRawRecordRunAsync(
        CsvRecordTextBatch batch, CancellationToken token)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_failed) throw new InvalidOperationException("The CSV reader cannot continue after a failed advance.");
        try
        {
            token.ThrowIfCancellationRequested();
            _options.CancellationToken.ThrowIfCancellationRequested();
            using var operation = LinkReadCancellation(token);
            var effective = operation?.Token ?? _options.CancellationToken;
            _current = Array.Empty<string>();
            CurrentPhysicalLineNumber = CurrentPhysicalEndLineNumber = null;
            var capture = await _records.CaptureRawRecordRunAsync(batch, effective).ConfigureAwait(false);
            effective.ThrowIfCancellationRequested();
            if (capture.Count > 0)
                SetCurrent(new BufferedRecord(_records.Current.Values, _records.StartLine, _records.EndLine));
            return capture;
        }
        catch { _failed = true; throw; }
    }

    private CancellationTokenSource? LinkReadCancellation(CancellationToken token) =>
        token.CanBeCanceled && token != _options.CancellationToken && token != _openingToken && token != _loadToken
            ? CancellationTokenSource.CreateLinkedTokenSource(token, _options.CancellationToken) : null;

    private void SetCurrent(BufferedRecord record)
    {
        _sourceValueCount = record.Values.Count;
        _current = CsvDocument.BuildParsedStringValues(record.Values, _headerCount, _options);
        CurrentPhysicalLineNumber = record.Start;
        CurrentPhysicalEndLineNumber = record.End;
    }

    public string GetString(int ordinal) => _current[ordinal];
    public bool HasStaticValues => _staticValues.Length > 0;
    public object? GetRawValue(int ordinal)
    {
        int sourceCount = _headerCount - _staticValues.Length;
        if (ordinal >= sourceCount) return _staticValues[ordinal - sourceCount];
        return IsNull(ordinal, _options.NullValue) ? null : GetString(ordinal);
    }
    public ReadOnlySpan<char> GetSpan(int ordinal) => _current is CsvIncrementalFieldValues fields
        ? fields.GetSpan(ordinal) : GetString(ordinal).AsSpan();
    internal ReadOnlySpan<char> GetSpan(int ordinal, out string? materialized)
    {
        if (_current is CsvIncrementalFieldValues fields)
        {
            materialized = fields.GetMaterializedString(ordinal);
            return fields.GetSpan(ordinal);
        }
        materialized = GetString(ordinal);
        return materialized.AsSpan();
    }
    public bool IsMissing(int ordinal) => ordinal >= _sourceValueCount && ordinal < _headerCount - (_options.StaticColumns?.Count ?? 0);
    public bool IsNull(int ordinal, string? nullValue) => !IsMissing(ordinal) && nullValue is not null &&
        GetSpan(ordinal).SequenceEqual(nullValue.AsSpan());
    public int CopyStringValues(object[] values, int count, string? nullValue)
    {
        int copied = Math.Min(count, _current.Count);
        for (int i = 0; i < copied; i++) values[i] = IsNull(i, nullValue) ? DBNull.Value : GetString(i);
        return copied;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try { _records.Dispose(); }
        finally { _lifetime.Dispose(); }
        _buffered.Clear();
        _current = Array.Empty<string>();
    }
}
#endif
