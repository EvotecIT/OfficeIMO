#nullable enable

#if NET8_0_OR_GREATER
using System.Buffers;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Data;

namespace OfficeIMO.CSV;

/// <summary>Creates one result from a transient, span-backed CSV record.</summary>
/// <remarks>
/// The callback can run concurrently. The record and spans obtained from it are valid only for the
/// duration of the callback and must not be retained.
/// </remarks>
public delegate T CsvRecordFactory<T>(CsvRecord record);

/// <summary>Builds a record factory after the CSV header has been resolved.</summary>
public delegate CsvRecordFactory<T> CsvRecordFactoryBuilder<T>(CsvRecordHeader header);

/// <summary>Resolved CSV header metadata for a transient-record projection.</summary>
public sealed class CsvRecordHeader
{
    private readonly string[] _names;
    private readonly Dictionary<string, int> _ordinals;

    internal CsvRecordHeader(CsvDataReader reader)
    {
        _names = new string[reader.FieldCount];
        _ordinals = new Dictionary<string, int>(reader.FieldCount, StringComparer.OrdinalIgnoreCase);
        for (int ordinal = 0; ordinal < _names.Length; ordinal++)
        {
            string name = reader.GetName(ordinal);
            _names[ordinal] = name;
            if (!_ordinals.ContainsKey(name))
            {
                _ordinals.Add(name, ordinal);
            }
        }
    }

    /// <summary>Gets the number of columns.</summary>
    public int Count => _names.Length;

    /// <summary>Gets a column name by ordinal.</summary>
    public string this[int ordinal] => _names[ordinal];

    /// <summary>Returns the ordinal for a column name.</summary>
    public int GetOrdinal(string name)
    {
        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        return _ordinals.TryGetValue(name, out int ordinal)
            ? ordinal
            : throw new IndexOutOfRangeException($"Column '{name}' was not found.");
    }

    /// <summary>Attempts to find a column ordinal by name.</summary>
    public bool TryGetOrdinal(string name, out int ordinal)
    {
        if (name is null)
        {
            ordinal = -1;
            return false;
        }

        return _ordinals.TryGetValue(name, out ordinal);
    }
}

/// <summary>A transient, span-backed CSV record passed to a projection callback.</summary>
public readonly ref struct CsvRecord
{
    private readonly CsvParser.CsvTextDataReaderBatch? _batch;
    private readonly CsvParser.CsvTextDataReaderRowSource? _source;
    private readonly CsvDataReader? _reader;

    internal CsvRecord(CsvParser.CsvTextDataReaderBatch batch)
    {
        _batch = batch;
        _source = null;
        _reader = null;
    }

    internal CsvRecord(CsvParser.CsvTextDataReaderRowSource source)
    {
        _batch = null;
        _source = source;
        _reader = null;
    }

    internal CsvRecord(CsvDataReader reader)
    {
        _batch = null;
        _source = null;
        _reader = reader;
    }

    /// <summary>Gets the number of source fields.</summary>
    public int FieldCount => _batch?.SourceColumnCount ?? _source?.SourceColumnCount ?? _reader!.FieldCount;

    /// <summary>Gets a transient unescaped field span.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public ReadOnlySpan<char> GetSpan(int ordinal) => _batch is not null
        ? _batch.GetSpan(ordinal)
        : _source is not null
            ? _source.GetSpan(ordinal)
            : _reader!.GetCurrentSourceString(ordinal).AsSpan();

    /// <summary>Materializes a field as a string.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public string GetString(int ordinal) => _batch is not null
        ? _batch.MaterializeString(ordinal)
        : _source is not null
            ? _source.GetString(ordinal)
            : _reader!.GetCurrentSourceString(ordinal);

    /// <summary>Returns whether the source row omitted the field at the specified ordinal.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool IsMissing(int ordinal) => _batch is not null
        ? _batch.IsMissing(ordinal)
        : _source is not null
            ? _source.IsMissing(ordinal)
            : _reader!.IsCurrentFieldMissing(ordinal);

    /// <summary>Returns whether the field matches the configured CSV null marker.</summary>
    /// <remarks>A missing field is not a configured null field; use <see cref="IsMissing"/> to distinguish it.</remarks>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool IsNull(int ordinal) => !IsMissing(ordinal) && (_batch is not null
        ? _batch.IsConfiguredNull(ordinal)
        : _source is not null
            ? _source.IsNull(ordinal, _source.Options.NullValue)
            : _reader!.IsDBNull(ordinal));

    /// <summary>Parses a Boolean field, accepting true/false and 0/1.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool GetBoolean(int ordinal)
    {
        if (_reader is not null)
        {
            return _reader.GetBoolean(ordinal);
        }
        ReadOnlySpan<char> value = GetSpan(ordinal);
        if (bool.TryParse(value, out bool result))
        {
            return result;
        }
        if (value.Length == 1 && (value[0] == '0' || value[0] == '1'))
        {
            return value[0] == '1';
        }
        throw CreateFormatException(ordinal, "Boolean");
    }

    /// <summary>Parses a 32-bit integer field using the reader culture.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public int GetInt32(int ordinal)
    {
        if (_reader is not null)
        {
            return _reader.GetInt32(ordinal);
        }
        ReadOnlySpan<char> value = GetSpan(ordinal);
        CultureInfo culture = _batch?.Culture ?? _source!.Options.Culture;
        if (ReferenceEquals(culture, CultureInfo.InvariantCulture) &&
            CsvDataProjectionConverter.TryParseInvariantInt32(value, out int result))
        {
            return result;
        }
        if (int.TryParse(value, NumberStyles.Any, culture, out result))
        {
            return result;
        }
        throw CreateFormatException(ordinal, "Int32");
    }

    /// <summary>Parses a 64-bit integer field using the reader culture.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public long GetInt64(int ordinal)
    {
        if (_reader is not null)
        {
            return _reader.GetInt64(ordinal);
        }
        CultureInfo culture = _batch?.Culture ?? _source!.Options.Culture;
        if (long.TryParse(GetSpan(ordinal), NumberStyles.Any, culture, out long result))
        {
            return result;
        }
        throw CreateFormatException(ordinal, "Int64");
    }

    /// <summary>Parses a decimal field using the reader culture.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public decimal GetDecimal(int ordinal)
    {
        if (_reader is not null)
        {
            return _reader.GetDecimal(ordinal);
        }
        ReadOnlySpan<char> value = GetSpan(ordinal);
        CultureInfo culture = _batch?.Culture ?? _source!.Options.Culture;
        if (ReferenceEquals(culture, CultureInfo.InvariantCulture) &&
            CsvDataProjectionConverter.TryParseInvariantDecimal(value, out decimal result))
        {
            return result;
        }
        if (decimal.TryParse(value, NumberStyles.Any, culture, out result))
        {
            return result;
        }
        throw CreateFormatException(ordinal, "Decimal");
    }

    /// <summary>Parses a double-precision number using the reader culture.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public double GetDouble(int ordinal)
    {
        if (_reader is not null)
        {
            return _reader.GetDouble(ordinal);
        }
        CultureInfo culture = _batch?.Culture ?? _source!.Options.Culture;
        if (double.TryParse(GetSpan(ordinal), NumberStyles.Any, culture, out double result))
        {
            return result;
        }
        throw CreateFormatException(ordinal, "Double");
    }

    /// <summary>Parses a DateTime field using configured CSV formats and round-trip semantics.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public DateTime GetDateTime(int ordinal)
    {
        if (_reader is not null)
        {
            return _reader.GetDateTime(ordinal);
        }
        CultureInfo culture = _batch?.Culture ?? _source!.Options.Culture;
        IReadOnlyList<string>? dateTimeFormats = _batch?.DateTimeFormats ?? _source!.Options.DateTimeFormats;
        if (CsvDataProjectionConverter.TryParseDateTime(
                GetSpan(ordinal),
                culture,
                dateTimeFormats,
                out DateTime result))
        {
            return result;
        }
        throw CreateFormatException(ordinal, "DateTime");
    }

    /// <summary>Parses a GUID field.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public Guid GetGuid(int ordinal)
    {
        if (_reader is not null)
        {
            return _reader.GetGuid(ordinal);
        }
        if (Guid.TryParse(GetSpan(ordinal), out Guid result))
        {
            return result;
        }
        throw CreateFormatException(ordinal, "Guid");
    }

    private FormatException CreateFormatException(int ordinal, string destinationType) =>
        new($"CSV field {ordinal} with value '{GetString(ordinal)}' cannot be converted to {destinationType}.");
}

public sealed partial class CsvDocument
{
    /// <summary>
    /// Reads decoded CSV text and projects records through bounded workers while preserving source order.
    /// </summary>
    /// <remarks>
    /// Header ordinals are resolved once by <paramref name="factoryBuilder"/>. The returned factory can
    /// run concurrently and must not mutate shared state. Eligible text is supplied as span-backed record
    /// batches by the AVX2 producer when available and by the scalar producer on other architectures.
    /// Unsupported records continue correctly on the calling thread. Enumeration owns and disposes the
    /// underlying reader.
    /// </remarks>
    /// <typeparam name="T">Result row type.</typeparam>
    /// <param name="text">Decoded CSV text.</param>
    /// <param name="factoryBuilder">Builds the concurrent record factory from resolved headers.</param>
    /// <param name="loadOptions">Optional CSV parsing settings.</param>
    /// <param name="readerOptions">Optional reader projection settings.</param>
    /// <param name="parallelOptions">Optional worker and batch settings.</param>
    /// <param name="cancellationToken">Cancels parsing and pending projection work.</param>
    public static IEnumerable<T> ReadTextRowsAsParallel<T>(
        string text,
        CsvRecordFactoryBuilder<T> factoryBuilder,
        CsvLoadOptions? loadOptions = null,
        CsvDataReaderOptions? readerOptions = null,
        ParallelRowMappingOptions? parallelOptions = null,
        CancellationToken cancellationToken = default)
    {
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (factoryBuilder is null) throw new ArgumentNullException(nameof(factoryBuilder));
        return EnumerateTextRowsAsParallel(
            text,
            factoryBuilder,
            loadOptions,
            readerOptions,
            parallelOptions,
            cancellationToken);
    }

    private static IEnumerable<T> EnumerateTextRowsAsParallel<T>(
        string text,
        CsvRecordFactoryBuilder<T> factoryBuilder,
        CsvLoadOptions? loadOptions,
        CsvDataReaderOptions? readerOptions,
        ParallelRowMappingOptions? parallelOptions,
        CancellationToken cancellationToken)
    {
        if (readerOptions?.ParallelProcessing is not null)
        {
            throw new ArgumentException(
                "ReadTextRowsAsParallel uses parallelOptions; CsvDataReaderOptions.ParallelProcessing must be omitted.",
                nameof(readerOptions));
        }

        using var reader = (CsvDataReader)OpenTextDataReader(text, loadOptions, readerOptions);
        CsvRecordFactory<T> factory = factoryBuilder(new CsvRecordHeader(reader))
            ?? throw new InvalidOperationException("The CSV record factory builder returned null.");
        ParallelRowMappingOptions options = parallelOptions ?? new ParallelRowMappingOptions();
        int degreeOfParallelism = options.GetDegreeOfParallelism();
        int batchSize = options.GetBatchSize(CsvParser.GetPreferredTextParallelBatchSize());
        if (degreeOfParallelism > 1 &&
            reader.TryPrepareTextPartitioning(
                cancellationToken,
                out CsvParser.CsvTextDataReaderRowSource? textSource,
                out int dataStart) &&
            TryCreateTextPartitions(
                text,
                dataStart,
                degreeOfParallelism,
                batchSize,
                textSource!.Options,
                cancellationToken,
                out CsvTextPartition[]? partitions))
        {
            foreach (T row in EnumeratePartitionedTextRows(
                         text,
                         textSource,
                         partitions!,
                         degreeOfParallelism,
                         batchSize,
                         factory,
                         cancellationToken))
            {
                yield return row;
            }
            yield break;
        }
        if (degreeOfParallelism == 1)
        {
            bool sequentialRemainderNeeded = false;
            while (true)
            {
                if (!reader.TryReadCsvRecordBatch(batchSize, cancellationToken, out CsvParser.CsvTextDataReaderBatch? batch))
                {
                    sequentialRemainderNeeded = true;
                    break;
                }
                if (batch is null)
                {
                    break;
                }

                using (batch)
                {
                    while (true)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (!batch.Read()) break;
                        cancellationToken.ThrowIfCancellationRequested();
                        yield return factory(new CsvRecord(batch));
                    }
                }
            }

            while (sequentialRemainderNeeded)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!reader.Read()) break;
                cancellationToken.ThrowIfCancellationRequested();
                yield return factory(new CsvRecord(reader));
            }
            yield break;
        }

        using var stop = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var pending = new Queue<Task<CsvMappedRecordBatch<T>>>(degreeOfParallelism);
        bool useSequentialRemainder = false;
        ExceptionDispatchInfo? sourceError = null;
        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                CsvParser.CsvTextDataReaderBatch? batch;
                try
                {
                    if (!reader.TryReadCsvRecordBatch(batchSize, cancellationToken, out batch))
                    {
                        useSequentialRemainder = true;
                        break;
                    }
                }
                catch (Exception exception) when (!cancellationToken.IsCancellationRequested)
                {
                    sourceError = ExceptionDispatchInfo.Capture(exception);
                    break;
                }
                if (batch is null) break;

                pending.Enqueue(Task.Factory.StartNew(
                    () => MapRecordBatch(batch, factory, stop.Token),
                    CancellationToken.None,
                    TaskCreationOptions.DenyChildAttach,
                    TaskScheduler.Default));

                if (pending.Count < degreeOfParallelism) continue;
                foreach (T row in AwaitRecordBatch(pending, stop, cancellationToken)) yield return row;
            }

            while (pending.Count > 0)
            {
                foreach (T row in AwaitRecordBatch(pending, stop, cancellationToken)) yield return row;
            }

            if (useSequentialRemainder)
            {
                while (true)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!reader.Read()) break;
                    cancellationToken.ThrowIfCancellationRequested();
                    yield return factory(new CsvRecord(reader));
                }
            }
            sourceError?.Throw();
        }
        finally
        {
            stop.Cancel();
            while (pending.Count > 0)
            {
                Task<CsvMappedRecordBatch<T>> task = pending.Dequeue();
                try
                {
                    task.GetAwaiter().GetResult().Return();
                }
                catch
                {
                    // Observe remaining canceled or faulted work after ordered propagation.
                }
            }
        }
    }

    private static IEnumerable<T> AwaitRecordBatch<T>(
        Queue<Task<CsvMappedRecordBatch<T>>> pending,
        CancellationTokenSource stop,
        CancellationToken cancellationToken)
    {
        Task<CsvMappedRecordBatch<T>> task = pending.Dequeue();
        CsvMappedRecordBatch<T> batch;
        try
        {
            batch = task.GetAwaiter().GetResult();
        }
        catch
        {
            stop.Cancel();
            throw;
        }

        try
        {
            for (int index = 0; index < batch.Count; index++)
            {
                if (cancellationToken.CanBeCanceled)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }
                yield return batch.Rows[index];
            }
        }
        finally
        {
            batch.Return();
        }
    }

    private static CsvMappedRecordBatch<T> MapRecordBatch<T>(
        CsvParser.CsvTextDataReaderBatch batch,
        CsvRecordFactory<T> factory,
        CancellationToken cancellationToken)
    {
        T[] rows = ArrayPool<T>.Shared.Rent(batch.RowCount);
        using (batch)
        {
            try
            {
                int index = 0;
                while (batch.Read())
                {
                    if ((index & 63) == 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                    }
                    rows[index++] = factory(new CsvRecord(batch));
                }
                cancellationToken.ThrowIfCancellationRequested();
                return new CsvMappedRecordBatch<T>(rows, index);
            }
            catch
            {
                ArrayPool<T>.Shared.Return(
                    rows,
                    clearArray: RuntimeHelpers.IsReferenceOrContainsReferences<T>());
                throw;
            }
        }
    }

    private readonly struct CsvMappedRecordBatch<T>
    {
        internal CsvMappedRecordBatch(T[] rows, int count)
        {
            Rows = rows;
            Count = count;
        }

        internal T[] Rows { get; }

        internal int Count { get; }

        internal void Return() => ArrayPool<T>.Shared.Return(
            Rows,
            clearArray: RuntimeHelpers.IsReferenceOrContainsReferences<T>());
    }

}
#endif
