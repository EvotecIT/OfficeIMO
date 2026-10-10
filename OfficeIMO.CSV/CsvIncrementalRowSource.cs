#nullable enable
#if NET8_0_OR_GREATER
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

internal sealed class CsvIncrementalRowSource : ICsvAsyncDataReaderRowSource, ICsvDataReaderPositionSource
{
    internal readonly record struct BufferedRecord(IReadOnlyList<string> Values, int Start, int End);
    private readonly CsvParser.IncrementalRecords _records;
    private readonly CsvLoadOptions _options;
    private readonly CancellationTokenSource _lifetime;
    private readonly CancellationToken _openingToken, _loadToken;
    private readonly int _headerCount;
    private readonly Queue<BufferedRecord> _buffered;
    private IReadOnlyList<string> _current = Array.Empty<string>();
    private bool _disposed;
    private bool _failed;
    private int _sourceValueCount;
    private readonly object?[] _staticValues;

    internal CsvIncrementalRowSource(CsvParser.IncrementalRecords records, CsvLoadOptions options,
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

    private async ValueTask<bool> AdvanceAsync(bool asynchronous, CancellationToken token)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_failed) throw new InvalidOperationException("The CSV reader cannot continue after a failed advance.");
        try { return await AdvanceCoreAsync(asynchronous, token).ConfigureAwait(false); }
        catch { _failed = true; throw; }
    }

    private async ValueTask<bool> AdvanceCoreAsync(bool asynchronous, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        _options.CancellationToken.ThrowIfCancellationRequested();
        // The lifetime already observes both tokens used to open the reader. Only a new
        // per-read token needs another link; aggregation reuses its opening token for every row.
        using var operation = token.CanBeCanceled && token != _options.CancellationToken &&
            token != _openingToken && token != _loadToken
            ? CancellationTokenSource.CreateLinkedTokenSource(token, _options.CancellationToken) : null;
        var effective = operation?.Token ?? _options.CancellationToken;
        if (_buffered.Count > 0)
        {
            SetCurrent(_buffered.Dequeue());
            return true;
        }
        if (!await _records.ReadAsync(asynchronous, effective, reuseValues: true).ConfigureAwait(false))
        {
            _current = Array.Empty<string>();
            CurrentPhysicalLineNumber = CurrentPhysicalEndLineNumber = null;
            return false;
        }
        effective.ThrowIfCancellationRequested();
        SetCurrent(new BufferedRecord(_records.Current.Values, _records.StartLine, _records.EndLine));
        return true;
    }

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
