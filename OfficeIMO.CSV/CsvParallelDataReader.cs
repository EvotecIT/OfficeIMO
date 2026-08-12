#nullable enable

using System.Buffers;
using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

/// <summary>
/// Preserves the single-consumer data-reader contract while typed CSV batches are projected by
/// bounded workers. The source reader remains owned by one calling thread.
/// </summary>
internal sealed class CsvParallelDataReader : DbDataReader,
    ICsvDataReaderMetadata,
    ICsvDataReaderPositionMetadata,
    IDataReaderMappingMetadata,
    IDataReaderMappingErrorMetadata,
    IDataReaderFastMappingValues
{
    private readonly CsvDataReader _source;
    private readonly IDataReaderParallelBatchSource _batchSource;
    private readonly int _degreeOfParallelism;
    private readonly int _batchSize;
    private readonly int _fieldCount;
    private readonly string[] _names;
    private readonly Queue<Task<CsvDataReaderRawBatch>> _pending;
    private readonly CancellationTokenSource _stop;
    private CsvDataReaderRawBatch? _currentBatch;
    private ExceptionDispatchInfo? _sourceError;
    private int _currentRow = -1;
    private long _recordNumber;
    private bool? _hasRows;
    private bool _sourceEnded;
    private bool _useRawBatches;
    private bool _closed;

    internal CsvParallelDataReader(CsvDataReader source, CsvDataReaderParallelOptions options)
    {
        _source = source ?? throw new ArgumentNullException(nameof(source));
        if (options is null) throw new ArgumentNullException(nameof(options));
        _degreeOfParallelism = options.GetDegreeOfParallelism();
        _batchSize = options.GetBatchSize(source.PreferredParallelProcessingBatchSize);
        _fieldCount = source.FieldCount;
        _names = new string[_fieldCount];
        for (int ordinal = 0; ordinal < _fieldCount; ordinal++)
        {
            _names[ordinal] = source.GetName(ordinal);
        }

        _batchSource = source;
        _useRawBatches = !_batchSource.CanReadParallelBatches;
        _pending = new Queue<Task<CsvDataReaderRawBatch>>(_degreeOfParallelism);
        _stop = source.ProcessingCancellationToken.CanBeCanceled
            ? CancellationTokenSource.CreateLinkedTokenSource(source.ProcessingCancellationToken)
            : new CancellationTokenSource();
    }

    internal static DbDataReader Apply(CsvDataReader source, CsvDataReaderOptions options)
    {
        CsvDataReaderParallelOptions? parallel = options.ParallelProcessing;
        if (parallel is null)
        {
            return source;
        }

        int degreeOfParallelism;
        try
        {
            degreeOfParallelism = parallel.GetDegreeOfParallelism();
            parallel.GetBatchSize(source.PreferredParallelProcessingBatchSize);
        }
        catch
        {
            source.Dispose();
            throw;
        }

        if (degreeOfParallelism == 1 || !source.CanBenefitFromParallelProcessing)
        {
            return source;
        }

        return new CsvParallelDataReader(source, parallel);
    }

    public override object this[int ordinal] => GetValue(ordinal);

    public override object this[string name] => GetValue(GetOrdinal(name));

    public override int Depth => 0;

    public char Delimiter => _source.Delimiter;

    public long RecordNumber => IsPositionedOnRow ? _recordNumber : 0;

    public int? PhysicalLineNumber => IsPositionedOnRow
        ? _currentBatch!.GetPhysicalLineNumber(_currentRow)
        : null;

    public int? PhysicalEndLineNumber => IsPositionedOnRow
        ? _currentBatch!.GetPhysicalEndLineNumber(_currentRow)
        : null;

    public override int FieldCount => _fieldCount;

    public override bool HasRows
    {
        get
        {
            if (_closed) return false;
            if (_hasRows.HasValue) return _hasRows.Value;
            _hasRows = EnsureCurrentBatch(CancellationToken.None);
            return _hasRows.Value;
        }
    }

    public override bool IsClosed => _closed;

    public override int RecordsAffected => -1;

    CultureInfo IDataReaderMappingMetadata.MappingCulture =>
        ((IDataReaderMappingMetadata)_source).MappingCulture;

    IReadOnlyList<string>? IDataReaderMappingMetadata.MappingDateTimeFormats =>
        ((IDataReaderMappingMetadata)_source).MappingDateTimeFormats;

    Func<object, Type, CultureInfo, (bool ok, object? value)>? IDataReaderMappingMetadata.MappingTypeConverter =>
        ((IDataReaderMappingMetadata)_source).MappingTypeConverter;

    bool IDataReaderMappingMetadata.RequireAllColumnsMapped =>
        ((IDataReaderMappingMetadata)_source).RequireAllColumnsMapped;

    DataMappingErrorValuePolicy IDataReaderMappingErrorMetadata.MappingErrorValuePolicy =>
        ((IDataReaderMappingErrorMetadata)_source).MappingErrorValuePolicy;

    bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => false;

    private bool IsPositionedOnRow => !_closed &&
        _currentBatch is not null &&
        (uint)_currentRow < (uint)_currentBatch.Count;

    public override bool Read() => ReadCore(CancellationToken.None);

    public override async Task<bool> ReadAsync(CancellationToken cancellationToken)
    {
        if (_closed) return false;
        cancellationToken.ThrowIfCancellationRequested();
        _stop.Token.ThrowIfCancellationRequested();
        if (!await EnsureCurrentBatchAsync(cancellationToken).ConfigureAwait(false))
        {
            _hasRows ??= false;
            _recordNumber = 0;
            return false;
        }

        _hasRows ??= true;
        _currentRow++;
        _recordNumber++;
        return true;
    }

    private bool ReadCore(CancellationToken cancellationToken)
    {
        if (_closed) return false;
        cancellationToken.ThrowIfCancellationRequested();
        _stop.Token.ThrowIfCancellationRequested();

        if (!EnsureCurrentBatch(cancellationToken))
        {
            _hasRows ??= false;
            _recordNumber = 0;
            return false;
        }

        _hasRows ??= true;
        _currentRow++;
        _recordNumber++;
        return true;
    }

    private bool EnsureCurrentBatch(CancellationToken cancellationToken)
    {
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            _stop.Token.ThrowIfCancellationRequested();
            if (_currentBatch is not null && _currentRow + 1 < _currentBatch.Count)
            {
                return true;
            }

            ReleaseCurrentBatchAndPropagateError();
            FillPending(cancellationToken);
            if (_pending.Count == 0)
            {
                _sourceError?.Throw();
                return false;
            }

            Task<CsvDataReaderRawBatch> task = _pending.Dequeue();
            try
            {
                _currentBatch = task.GetAwaiter().GetResult();
            }
            catch
            {
                _stop.Cancel();
                throw;
            }

            _currentRow = -1;
            if (_currentBatch.Count != 0)
            {
                return true;
            }
        }
    }

    private async Task<bool> EnsureCurrentBatchAsync(CancellationToken cancellationToken)
    {
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            _stop.Token.ThrowIfCancellationRequested();
            if (_currentBatch is not null && _currentRow + 1 < _currentBatch.Count)
            {
                return true;
            }

            ReleaseCurrentBatchAndPropagateError();
            FillPending(cancellationToken);
            if (_pending.Count == 0)
            {
                _sourceError?.Throw();
                return false;
            }

            Task<CsvDataReaderRawBatch> task = _pending.Dequeue();
            using CancellationTokenRegistration registration = cancellationToken.Register(
                static state => ((CancellationTokenSource)state!).Cancel(),
                _stop);
            try
            {
                _currentBatch = await task.ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                _stop.Token.ThrowIfCancellationRequested();
            }
            catch
            {
                _stop.Cancel();
                throw;
            }

            _currentRow = -1;
            if (_currentBatch.Count != 0)
            {
                return true;
            }
        }
    }

    private void FillPending(CancellationToken cancellationToken)
    {
        while (!_sourceEnded && _sourceError is null && _pending.Count < _degreeOfParallelism)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                if (!_useRawBatches)
                {
                    if (!_batchSource.TryReadParallelBatch(
                            _batchSize,
                            cancellationToken,
                            out DbDataReader? detachedReader))
                    {
                        _useRawBatches = true;
                        continue;
                    }

                    if (detachedReader is null)
                    {
                        _sourceEnded = true;
                        break;
                    }

                    var csvBatchReader = (CsvDataReader)detachedReader;
                    int expectedRows = detachedReader is IDataReaderParallelBatchInfo info
                        ? info.ParallelBatchRowCount
                        : _batchSize;
                    _pending.Enqueue(Task.Factory.StartNew(
                        () => MaterializeDetachedBatch(csvBatchReader, expectedRows, _stop.Token),
                        CancellationToken.None,
                        TaskCreationOptions.DenyChildAttach,
                        TaskScheduler.Default));
                    continue;
                }

                CsvDataReaderRawBatch rawBatch = _source.ReadRawBatch(
                    _batchSize,
                    cancellationToken,
                    out bool reachedEnd);
                _sourceEnded = reachedEnd;
                if (rawBatch.Count == 0 && rawBatch.Error is null)
                {
                    rawBatch.Dispose();
                    break;
                }

                _pending.Enqueue(Task.Factory.StartNew(
                    () => _source.ConvertRawBatch(rawBatch, _stop.Token),
                    CancellationToken.None,
                    TaskCreationOptions.DenyChildAttach,
                    TaskScheduler.Default));
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception)
            {
                _sourceError = ExceptionDispatchInfo.Capture(exception);
            }
        }
    }

    private CsvDataReaderRawBatch MaterializeDetachedBatch(
        CsvDataReader batchReader,
        int expectedRows,
        CancellationToken cancellationToken)
    {
        var batch = new CsvDataReaderRawBatch(expectedRows, _fieldCount, includePositions: true);
        try
        {
            while (batch.Count < batch.RowCapacity)
            {
                if ((batch.Count & 63) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }

                bool hasRow;
                try
                {
                    hasRow = batchReader.Read();
                }
                catch (Exception exception) when (!(exception is OperationCanceledException))
                {
                    batch.SetError(batch.Count, exception);
                    break;
                }

                if (!hasRow) break;
                int offset = batch.Count * _fieldCount;
                try
                {
                    for (int ordinal = 0; ordinal < _fieldCount; ordinal++)
                    {
                        batch.Values[offset + ordinal] = batchReader.GetValue(ordinal);
                    }
                }
                catch (Exception exception) when (!(exception is OperationCanceledException))
                {
                    batch.SetError(batch.Count, exception);
                    break;
                }

                batch.SetPosition(
                    batch.Count,
                    batchReader.PhysicalLineNumber,
                    batchReader.PhysicalEndLineNumber);
                batch.Count++;
            }

            cancellationToken.ThrowIfCancellationRequested();
            return batch;
        }
        catch
        {
            batch.Dispose();
            throw;
        }
        finally
        {
            batchReader.Dispose();
        }
    }

    private void ReleaseCurrentBatchAndPropagateError()
    {
        if (_currentBatch is null) return;
        ExceptionDispatchInfo? error = _currentBatch.Error;
        _currentBatch.Dispose();
        _currentBatch = null;
        _currentRow = -1;
        error?.Throw();
    }

    public override bool GetBoolean(int ordinal) => (bool)GetValue(ordinal);

    public override byte GetByte(int ordinal) => (byte)GetValue(ordinal);

    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
        throw new NotSupportedException("CSV fields are exposed as scalar values.");

    public override char GetChar(int ordinal)
    {
        object value = GetValue(ordinal);
        if (value is char character) return character;
        if (value is string { Length: 1 } text) return text[0];
        throw new InvalidCastException("The CSV field does not contain one character.");
    }

    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length)
    {
        string value = Convert.ToString(GetValue(ordinal), MappingCulture) ?? string.Empty;
        if (buffer is null) return value.Length;
        if (dataOffset >= value.Length || length == 0) return 0;
        int offset = checked((int)dataOffset);
        int count = Math.Min(length, value.Length - offset);
        if (count <= 0) return 0;
        value.CopyTo(offset, buffer, bufferOffset, count);
        return count;
    }

    public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

    public override DateTime GetDateTime(int ordinal) => (DateTime)GetValue(ordinal);

    public override decimal GetDecimal(int ordinal) => (decimal)GetValue(ordinal);

    public override double GetDouble(int ordinal) => (double)GetValue(ordinal);

    public override IEnumerator GetEnumerator()
    {
        while (Read()) yield return this;
    }

    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
    public override Type GetFieldType(int ordinal) => _source.GetFieldType(ordinal);

    public override float GetFloat(int ordinal) => (float)GetValue(ordinal);

    public override Guid GetGuid(int ordinal) => (Guid)GetValue(ordinal);

    public override short GetInt16(int ordinal) => (short)GetValue(ordinal);

    public override int GetInt32(int ordinal) => (int)GetValue(ordinal);

    public override long GetInt64(int ordinal) => (long)GetValue(ordinal);

    public override string GetName(int ordinal) => _names[ordinal];

    public override int GetOrdinal(string name)
    {
        for (int ordinal = 0; ordinal < _names.Length; ordinal++)
        {
            if (string.Equals(_names[ordinal], name, StringComparison.OrdinalIgnoreCase)) return ordinal;
        }

        throw new IndexOutOfRangeException(name);
    }

    public override string GetString(int ordinal) => (string)GetValue(ordinal);

    public override object GetValue(int ordinal)
    {
        EnsureOpenRow();
        if ((uint)ordinal >= (uint)_fieldCount) throw new IndexOutOfRangeException();
        return _currentBatch!.Values[(_currentRow * _fieldCount) + ordinal] ?? DBNull.Value;
    }

    public override int GetValues(object[] values)
    {
        if (values is null) throw new ArgumentNullException(nameof(values));
        EnsureOpenRow();
        int count = Math.Min(values.Length, _fieldCount);
        Array.Copy(_currentBatch!.Values, _currentRow * _fieldCount, values, 0, count);
        return count;
    }

    public override bool IsDBNull(int ordinal) => ReferenceEquals(GetValue(ordinal), DBNull.Value);

    public override bool NextResult() => false;

    public override DataTable GetSchemaTable() => _source.GetSchemaTable();

    public override void Close()
    {
        if (_closed) return;
        _closed = true;
        _stop.Cancel();
        _currentBatch?.Dispose();
        _currentBatch = null;
        while (_pending.Count > 0)
        {
            try
            {
                _pending.Dequeue().GetAwaiter().GetResult().Dispose();
            }
            catch
            {
                // Dispose observes canceled and faulted workers after stopping new source work.
            }
        }

        _source.Dispose();
        _stop.Dispose();
        _recordNumber = 0;
        _currentRow = -1;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) Close();
        base.Dispose(disposing);
    }

    private CultureInfo MappingCulture =>
        ((IDataReaderMappingMetadata)_source).MappingCulture;

    private void EnsureOpenRow()
    {
        if (_closed) throw new InvalidOperationException("The reader is closed.");
        if (!IsPositionedOnRow) throw new InvalidOperationException("The reader is not positioned on a row.");
    }
}

internal sealed class CsvDataReaderRawBatch : IDisposable
{
    private readonly int?[]? _physicalLineNumbers;
    private readonly int?[]? _physicalEndLineNumbers;
    private bool _disposed;

    internal CsvDataReaderRawBatch(int rowCapacity, int fieldCount, bool includePositions)
    {
        RowCapacity = GetBoundedRowCapacity(rowCapacity, fieldCount);
        FieldCount = fieldCount;
        Values = ArrayPool<object?>.Shared.Rent(checked(RowCapacity * fieldCount));
        if (includePositions)
        {
            _physicalLineNumbers = ArrayPool<int?>.Shared.Rent(RowCapacity);
            _physicalEndLineNumbers = ArrayPool<int?>.Shared.Rent(RowCapacity);
        }
    }

    internal object?[] Values { get; }

    internal int RowCapacity { get; }

    internal int FieldCount { get; }

    internal int Count { get; set; }

    internal int FirstRecordIndex { get; set; }

    internal ExceptionDispatchInfo? Error { get; private set; }

    internal static int GetBoundedRowCapacity(int preferredBatchSize, int fieldCount)
    {
        if (preferredBatchSize <= 0) throw new ArgumentOutOfRangeException(nameof(preferredBatchSize));
        if (fieldCount <= 0) return 1;
        const int maximumElements = 4 * 1024 * 1024;
        return Math.Max(1, Math.Min(preferredBatchSize, maximumElements / fieldCount));
    }

    internal void SetError(int completedRowCount, Exception exception)
    {
        Count = completedRowCount;
        Error = ExceptionDispatchInfo.Capture(exception);
    }

    internal void SetPosition(int row, int? physicalLineNumber, int? physicalEndLineNumber)
    {
        if (_physicalLineNumbers is null || _physicalEndLineNumbers is null) return;
        _physicalLineNumbers[row] = physicalLineNumber;
        _physicalEndLineNumbers[row] = physicalEndLineNumber;
    }

    internal int? GetPhysicalLineNumber(int row) => _physicalLineNumbers?[row];

    internal int? GetPhysicalEndLineNumber(int row) => _physicalEndLineNumbers?[row];

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        ArrayPool<object?>.Shared.Return(Values, clearArray: true);
        if (_physicalLineNumbers is not null)
        {
            Array.Clear(_physicalLineNumbers, 0, Count);
            ArrayPool<int?>.Shared.Return(_physicalLineNumbers);
        }

        if (_physicalEndLineNumbers is not null)
        {
            Array.Clear(_physicalEndLineNumbers, 0, Count);
            ArrayPool<int?>.Shared.Return(_physicalEndLineNumbers);
        }
    }
}
