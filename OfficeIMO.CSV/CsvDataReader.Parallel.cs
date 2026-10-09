#nullable enable

using System.Collections;
using System.Buffers;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using CsvDataReaderTextRowSource = OfficeIMO.CSV.ICsvDataReaderTextRowSource;

namespace OfficeIMO.CSV;

internal sealed partial class CsvDataReader
{
    internal bool HasIncrementalSource =>
#if NET8_0_OR_GREATER
        _textRowSource is ICsvAsyncDataReaderRowSource;
#else
        false;
#endif

#if NET8_0_OR_GREATER
    bool IDataReaderParallelBatchSource.CanReadParallelBatches =>
        !_closed && !_checkedForRows && _currentRawRow is null &&
        _currentStringRow is null && !_hasCurrentTextRow &&
        _textRowSource is CsvParser.CsvTextDataReaderRowSource { CanTakeParallelBatch: true };

    int IDataReaderParallelBatchSource.PreferredParallelBatchSize =>
        (_textRowSource as CsvParser.CsvTextDataReaderRowSource)?.PreferredParallelBatchSize ?? 128;
#else
    bool IDataReaderParallelBatchSource.CanReadParallelBatches => false;

    int IDataReaderParallelBatchSource.PreferredParallelBatchSize => 128;
#endif

    bool IDataReaderFastMappingValues.HasOnlyNonNullFastValues => _useDirectTextSourceStrings;

    int IDataReaderParallelBatchInfo.ParallelBatchRowCount =>
        (_textRowSource as ICsvDataReaderParallelBatchInfo)?.RowCount ?? 0;

    internal bool CanBenefitFromParallelProcessing =>
        !_closed && _columns.Length != 0 && !_useRawStringValues;

    internal int PreferredParallelProcessingBatchSize =>
#if NET8_0_OR_GREATER
        _textRowSource is CsvParser.CsvTextDataReaderRowSource ? 4096 : 256;
#else
        256;
#endif

    internal CancellationToken ProcessingCancellationToken => _processingCancellationToken;

    bool IDataReaderParallelBatchSource.TryReadParallelBatch(
        int preferredBatchSize,
        CancellationToken cancellationToken,
        out DbDataReader? batchReader)
    {
        cancellationToken.ThrowIfCancellationRequested();
        batchReader = null;
#if NET8_0_OR_GREATER
        if (_closed || _checkedForRows || _currentRawRow is not null ||
            _currentStringRow is not null || _hasCurrentTextRow ||
            _textRowSource is not CsvParser.CsvTextDataReaderRowSource textRows)
        {
            return false;
        }

        if (!textRows.TryTakeParallelBatch(
                preferredBatchSize,
                cancellationToken,
                out ICsvDataReaderTextRowSource? batchRows))
        {
            return false;
        }

        if (batchRows is null)
        {
            return true;
        }

        int firstRowIndex = _rowIndex;
        int batchRowCount = (batchRows as ICsvDataReaderParallelBatchInfo)?.RowCount ?? 0;
        _rowIndex = checked(_rowIndex + batchRowCount);
        batchReader = new CsvDataReader(
            _columns,
            batchRows,
            _sourceColumnCount,
            _stringRowOptions!,
            _culture,
            _dateTimeFormats,
            initialRowIndex: firstRowIndex);
        return true;
#else
        return false;
#endif
    }

#if NET8_0_OR_GREATER
    internal bool TryPrepareTextPartitioning(
        CancellationToken cancellationToken,
        out CsvParser.CsvTextDataReaderRowSource? source,
        out int dataStart)
    {
        source = null;
        dataStart = 0;
        cancellationToken.ThrowIfCancellationRequested();
        if (!_useRawStringValues || _closed || _checkedForRows || _currentRawRow is not null ||
            _currentStringRow is not null || _hasCurrentTextRow ||
            _textRowSource is not CsvParser.CsvTextDataReaderRowSource textRows ||
            !textRows.CanTakeParallelBatch)
        {
            return false;
        }

        dataStart = textRows.PrepareForParallelPartition(cancellationToken);
        source = textRows;
        return true;
    }

    internal bool TryReadCsvRecordBatch(
        int preferredBatchSize,
        CancellationToken cancellationToken,
        out CsvParser.CsvTextDataReaderBatch? batch)
    {
        cancellationToken.ThrowIfCancellationRequested();
        batch = null;
        if (!_useRawStringValues || _closed || _checkedForRows || _currentRawRow is not null ||
            _currentStringRow is not null || _hasCurrentTextRow ||
            _textRowSource is not CsvParser.CsvTextDataReaderRowSource textRows)
        {
            return false;
        }

        if (!textRows.TryTakeParallelBatch(
                preferredBatchSize,
                cancellationToken,
                out ICsvDataReaderTextRowSource? batchRows))
        {
            return false;
        }

        batch = batchRows as CsvParser.CsvTextDataReaderBatch;
        batch?.SetNullValue(_stringNullValue);
        return batchRows is null || batch is not null;
    }

    internal bool IsCurrentFieldMissing(int ordinal)
    {
        EnsureOpenRow();
        if ((uint)ordinal >= (uint)_columns.Length)
        {
            throw new IndexOutOfRangeException();
        }

        if (_textRowSource is not null)
        {
            return _textRowSource.IsMissing(ordinal);
        }
        if (_currentStringRow is not null)
        {
            return ordinal < _sourceColumnCount && ordinal >= _currentStringRow.Count;
        }
        return ordinal >= _currentRawRow!.Length;
    }

    internal string GetCurrentSourceString(int ordinal)
    {
        EnsureOpenRow();
        if ((uint)ordinal >= (uint)_columns.Length)
        {
            throw new IndexOutOfRangeException();
        }

        if (_textRowSource is not null)
        {
            return _textRowSource.GetString(ordinal);
        }
        if (_currentStringRow is not null)
        {
            if (ordinal >= _sourceColumnCount)
            {
                return Convert.ToString(GetRawValue(ordinal), _culture) ?? string.Empty;
            }
            return ordinal < _currentStringRow.Count ? _currentStringRow[ordinal] : string.Empty;
        }

        if (_rawRowsAreParsedStringsOnly)
        {
            object? value = ordinal < _currentRawRow!.Length ? _currentRawRow[ordinal] : null;
            return value as string ?? string.Empty;
        }

        throw new InvalidOperationException(
            "The current CSV row is not backed by decoded source text.");
    }
#endif

    internal CsvDataReaderRawBatch ReadRawBatch(
        int preferredBatchSize,
        CancellationToken cancellationToken,
        out bool reachedEnd)
    {
        var capture = ReadRawBatchAsync(preferredBatchSize, asynchronous: false, cancellationToken).GetAwaiter().GetResult();
        reachedEnd = capture.ReachedEnd;
        return capture.Batch;
    }

    internal async Task<(CsvDataReaderRawBatch Batch, bool ReachedEnd)> ReadRawBatchAsync(
        int preferredBatchSize, bool asynchronous, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        int rowCapacity = CsvDataReaderRawBatch.GetBoundedRowCapacity(preferredBatchSize, _columns.Length);
        var batch = new CsvDataReaderRawBatch(rowCapacity, _columns.Length, includePositions: true);
        bool reachedEnd = false;
        try
        {
            while (batch.Count < rowCapacity)
            {
                if ((batch.Count & 63) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }

                bool hasRow;
                try
                {
                    hasRow = asynchronous
                        ? await ReadAsync(cancellationToken).ConfigureAwait(false)
                        : ReadCore(cancellationToken);
                }
                catch (Exception exception) when (!(exception is OperationCanceledException))
                {
                    batch.SetError(batch.Count, exception);
                    reachedEnd = true;
                    break;
                }

                if (!hasRow)
                {
                    reachedEnd = true;
                    break;
                }

                if (batch.Count == 0)
                {
                    batch.FirstRecordIndex = _rowIndex;
                }

                int offset = batch.Count * _columns.Length;
                try
                {
                    for (int ordinal = 0; ordinal < _columns.Length; ordinal++)
                    {
                        batch.Values[offset + ordinal] = GetRawValue(ordinal);
                    }
                }
                catch (Exception exception) when (!(exception is OperationCanceledException))
                {
                    batch.SetError(batch.Count, exception);
                    reachedEnd = true;
                    break;
                }

                batch.SetPosition(batch.Count, PhysicalLineNumber, PhysicalEndLineNumber);
                batch.Count++;
            }

            cancellationToken.ThrowIfCancellationRequested();
            return (batch, reachedEnd);
        }
        catch
        {
            batch.Dispose();
            throw;
        }
    }

    internal CsvDataReaderRawBatch ConvertRawBatch(
        CsvDataReaderRawBatch batch,
        CancellationToken cancellationToken)
    {
        try
        {
            for (int row = 0; row < batch.Count; row++)
            {
                if ((row & 63) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }

                int offset = row * _columns.Length;
                try
                {
                    for (int ordinal = 0; ordinal < _columns.Length; ordinal++)
                    {
                        batch.Values[offset + ordinal] = CsvDataProjectionConverter.ConvertValue(
                            batch.Values[offset + ordinal],
                            _columns[ordinal],
                            batch.FirstRecordIndex + row,
                            _culture,
                            _dateTimeFormats,
                            _mappingErrorValuePolicy);
                    }
                }
                catch (Exception exception) when (!(exception is OperationCanceledException))
                {
                    batch.SetError(row, exception);
                    break;
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            return batch;
        }
        catch
        {
            batch.Dispose();
            throw;
        }
    }

}
