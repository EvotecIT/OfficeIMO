#nullable enable

#if NET8_0_OR_GREATER
using System.Buffers;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    private static IEnumerable<T> EnumeratePartitionedTextRows<T>(
        string text,
        CsvParser.CsvTextDataReaderRowSource preparedSource,
        CsvTextPartition[] partitions,
        int degreeOfParallelism,
        int batchSize,
        CsvRecordFactory<T> factory,
        CancellationToken cancellationToken)
    {
        for (int waveStart = 0; waveStart < partitions.Length; waveStart += degreeOfParallelism)
        {
            int waveCount = Math.Min(degreeOfParallelism, partitions.Length - waveStart);
            var mapped = new CsvMappedRecordBatch<T>[waveCount];
            var completed = new bool[waveCount];
            try
            {
                try
                {
                    int capturedWaveStart = waveStart;
                    Parallel.For(
                        0,
                        waveCount,
                        new ParallelOptions {
                            MaxDegreeOfParallelism = waveCount,
                            CancellationToken = cancellationToken
                        },
                        index => {
                            mapped[index] = MapTextPartition(
                                text,
                                preparedSource.Options,
                                preparedSource.SourceColumnCount,
                                partitions[capturedWaveStart + index],
                                factory,
                                cancellationToken);
                            Volatile.Write(ref completed[index], true);
                        });
                }
                catch (AggregateException exception)
                {
                    IReadOnlyCollection<Exception> failures = exception.Flatten().InnerExceptions;
                    Exception first = failures.FirstOrDefault(
                        static failure => failure is not OperationCanceledException)
                        ?? failures.First();
                    ExceptionDispatchInfo.Capture(first).Throw();
                }
                for (int partitionIndex = 0; partitionIndex < mapped.Length; partitionIndex++)
                {
                    CsvMappedRecordBatch<T> batch = mapped[partitionIndex];
                    for (int rowIndex = 0; rowIndex < batch.Count; rowIndex++)
                    {
                        if (cancellationToken.CanBeCanceled)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                        }
                        yield return batch.Rows[rowIndex];
                    }
                }
            }
            finally
            {
                for (int index = 0; index < mapped.Length; index++)
                {
                    if (Volatile.Read(ref completed[index]))
                    {
                        mapped[index].Return();
                    }
                }
            }
        }
    }

    private static CsvMappedRecordBatch<T> MapTextPartition<T>(
        string text,
        CsvLoadOptions options,
        int sourceColumnCount,
        CsvTextPartition partition,
        CsvRecordFactory<T> factory,
        CancellationToken cancellationToken)
    {
        T[] rows = ArrayPool<T>.Shared.Rent(partition.RowCount);
        int outputIndex = 0;
        using var source = new CsvParser.CsvTextDataReaderRowSource(
            text,
            options,
            recordsToSkip: 0,
            sourceColumnCount,
            partition.Start,
            partition.End);
        try
        {
            while (source.Read())
            {
                if (cancellationToken.CanBeCanceled && (outputIndex & 63) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }
                if ((uint)outputIndex >= (uint)partition.RowCount)
                {
                    throw new InvalidDataException(
                        "Partitioned CSV record counting did not match parser output.");
                }
                rows[outputIndex++] = factory(new CsvRecord(source));
            }

            cancellationToken.ThrowIfCancellationRequested();

            if (outputIndex != partition.RowCount)
            {
                throw new InvalidDataException(
                    "Partitioned CSV record counting did not match parser output.");
            }

            return new CsvMappedRecordBatch<T>(rows, outputIndex);
        }
        catch
        {
            ArrayPool<T>.Shared.Return(
                rows,
                clearArray: RuntimeHelpers.IsReferenceOrContainsReferences<T>());
            throw;
        }
    }

    private static bool TryCreateTextPartitions(
        string text,
        int dataStart,
        int degreeOfParallelism,
        int batchSize,
        CsvLoadOptions options,
        CancellationToken cancellationToken,
        out CsvTextPartition[]? partitions)
    {
        partitions = null;
        if (dataStart >= text.Length ||
            degreeOfParallelism <= 1 ||
            batchSize <= 0 ||
            options.TrimWhitespace ||
            options.SkipCommentRows ||
            options.ProgressReportInterval != 0 ||
            options.ProgressCallback is not null ||
            options.CollectParseErrors ||
            options.ParseErrorAction != CsvParseErrorAction.Throw ||
            options.Delimiter is '\r' or '\n' or '"' ||
            !string.IsNullOrEmpty(options.DelimiterText) ||
            options.DetectDelimiter ||
            options.StaticColumns is not null ||
            options.MaxFieldLength is not null ||
            options.MaxQuotedFieldLength is not null ||
            options.NormalizeQuotes ||
            options.InternStrings)
        {
            return false;
        }

        var state = new CsvTextPartitionScanState(
            dataStart,
            text.Length,
            Math.Min(batchSize, 256));
        bool canCancel = cancellationToken.CanBeCanceled || options.CancellationToken.CanBeCanceled;
        int index = dataStart;
        int skipThrough = -1;
        if (Avx2.IsSupported)
        {
            ref ushort textStart = ref Unsafe.As<char, ushort>(
                ref MemoryMarshal.GetReference(text.AsSpan()));
            Vector256<ushort> quote = Vector256.Create((ushort)'"');
            Vector256<ushort> carriageReturn = Vector256.Create((ushort)'\r');
            Vector256<ushort> lineFeed = Vector256.Create((ushort)'\n');
            int vectorEnd = text.Length - Vector256<ushort>.Count;
            while (index <= vectorEnd)
            {
                if (canCancel && ((index - dataStart) & 0xFFFF) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    options.CancellationToken.ThrowIfCancellationRequested();
                }

                Vector256<ushort> values = Vector256.LoadUnsafe(
                    ref textStart,
                    (nuint)index);
                Vector256<ushort> matches = Avx2.Or(
                    Avx2.CompareEqual(values, quote),
                    Avx2.Or(
                        Avx2.CompareEqual(values, carriageReturn),
                        Avx2.CompareEqual(values, lineFeed)));
                uint mask = unchecked((uint)Avx2.MoveMask(matches.AsByte())) & 0x55555555u;
                while (mask != 0)
                {
                    int offset = BitOperations.TrailingZeroCount(mask) >> 1;
                    int specialIndex = index + offset;
                    mask &= mask - 1;
                    if (specialIndex <= skipThrough)
                    {
                        continue;
                    }
                    if (!TryProcessTextPartitionSpecial(
                            text,
                            specialIndex,
                            options,
                            ref state,
                            ref skipThrough))
                    {
                        return false;
                    }
                }
                index += Vector256<ushort>.Count;
            }
        }

        for (; index < text.Length; index++)
        {
            char value = text[index];
            if (value is not '"' and not '\r' and not '\n' || index <= skipThrough)
            {
                continue;
            }
            if (!TryProcessTextPartitionSpecial(
                    text,
                    index,
                    options,
                    ref state,
                    ref skipThrough))
            {
                return false;
            }
        }

        if (state.InQuotes || text.Length - state.RecordStart > ushort.MaxValue)
        {
            return false;
        }
        int finalLength = text.Length - state.RecordStart;
        if (finalLength > 0)
        {
            if (StartsWithSpecialComment(text, state.RecordStart, options))
            {
                return false;
            }
            state.PartitionRows++;
        }
        if (state.PartitionRows > 0)
        {
            state.Result.Add(new CsvTextPartition(
                state.PartitionStart,
                text.Length,
                state.PartitionRows));
        }

        if (state.Result.Count <= 1)
        {
            return false;
        }

        partitions = CreateBalancedTextPartitions(
            state.Result,
            degreeOfParallelism,
            batchSize);
        return true;
    }

    private static CsvTextPartition[] CreateBalancedTextPartitions(
        List<CsvTextPartition> chunks,
        int degreeOfParallelism,
        int batchSize)
    {
        long configuredWaveRows = (long)degreeOfParallelism * batchSize;
        int maximumWaveRows = configuredWaveRows >= int.MaxValue
            ? int.MaxValue
            : Math.Max(degreeOfParallelism, (int)configuredWaveRows);
        var result = new List<CsvTextPartition>(chunks.Count);
        int waveStart = 0;
        while (waveStart < chunks.Count)
        {
            int waveEnd = waveStart;
            int waveRows = 0;
            while (waveEnd < chunks.Count)
            {
                int nextRows = chunks[waveEnd].RowCount;
                if (waveRows != 0 && nextRows > maximumWaveRows - waveRows)
                {
                    break;
                }
                waveRows += nextRows;
                waveEnd++;
            }

            int workerCount = Math.Min(degreeOfParallelism, waveEnd - waveStart);
            int chunkIndex = waveStart;
            int remainingRows = waveRows;
            for (int worker = 0; worker < workerCount; worker++)
            {
                int workersRemaining = workerCount - worker;
                int targetRows = (remainingRows + workersRemaining - 1) / workersRemaining;
                int partitionStart = chunks[chunkIndex].Start;
                int partitionEnd = partitionStart;
                int partitionRows = 0;
                while (chunkIndex < waveEnd)
                {
                    int chunksAfterThis = waveEnd - chunkIndex - 1;
                    if (partitionRows != 0 && chunksAfterThis < workersRemaining - 1)
                    {
                        break;
                    }

                    CsvTextPartition next = chunks[chunkIndex];
                    if (partitionRows != 0 &&
                        Math.Abs(partitionRows - targetRows) <=
                        Math.Abs(partitionRows + next.RowCount - targetRows))
                    {
                        break;
                    }

                    partitionRows += next.RowCount;
                    partitionEnd = next.End;
                    chunkIndex++;
                }

                result.Add(new CsvTextPartition(
                    partitionStart,
                    partitionEnd,
                    partitionRows));
                remainingRows -= partitionRows;
            }

            waveStart = waveEnd;
        }

        return result.ToArray();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool TryProcessTextPartitionSpecial(
        string text,
        int index,
        CsvLoadOptions options,
        ref CsvTextPartitionScanState state,
        ref int skipThrough)
    {
        char current = text[index];
        if (current == '"')
        {
            if (state.InQuotes)
            {
                if (index + 1 < text.Length && text[index + 1] == '"')
                {
                    skipThrough = index + 1;
                    return true;
                }
                if (index - state.QuoteStart > ushort.MaxValue)
                {
                    return false;
                }
                int afterQuote = index + 1;
                if (afterQuote < text.Length &&
                    text[afterQuote] != options.Delimiter &&
                    text[afterQuote] is not '\r' and not '\n')
                {
                    return false;
                }
                state.InQuotes = false;
                return true;
            }

            if (index != state.RecordStart && text[index - 1] != options.Delimiter)
            {
                return false;
            }
            state.InQuotes = true;
            state.QuoteStart = index;
            return true;
        }

        if (state.InQuotes)
        {
            return true;
        }
        if (index - state.RecordStart > ushort.MaxValue ||
            StartsWithSpecialComment(text, state.RecordStart, options))
        {
            return false;
        }

        int recordEnd = index + 1;
        if (current == '\r' && recordEnd < text.Length && text[recordEnd] == '\n')
        {
            recordEnd++;
            skipThrough = recordEnd - 1;
        }
        if (index > state.RecordStart || options.AllowEmptyLines)
        {
            state.PartitionRows++;
        }
        state.RecordStart = recordEnd;
        state.QuoteStart = -1;
        if (state.PartitionRows == state.BatchSize)
        {
            state.Result.Add(new CsvTextPartition(
                state.PartitionStart,
                recordEnd,
                state.PartitionRows));
            state.PartitionStart = recordEnd;
            state.PartitionRows = 0;
        }
        return true;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool StartsWithSpecialComment(
        string text,
        int recordStart,
        CsvLoadOptions options) =>
        recordStart < text.Length &&
        text[recordStart] == options.CommentCharacter &&
        (options.SkipCommentRowsBeforeHeader || options.RecognizeW3CFieldsHeader);

    private struct CsvTextPartitionScanState
    {
        internal CsvTextPartitionScanState(
            int dataStart,
            int textLength,
            int batchSize)
        {
            Result = new List<CsvTextPartition>(Math.Max(4, (textLength - dataStart) / 262_144));
            PartitionStart = dataStart;
            PartitionRows = 0;
            BatchSize = batchSize;
            RecordStart = dataStart;
            QuoteStart = -1;
            InQuotes = false;
        }

        internal List<CsvTextPartition> Result;
        internal int PartitionStart;
        internal int PartitionRows;
        internal int BatchSize;
        internal int RecordStart;
        internal int QuoteStart;
        internal bool InQuotes;
    }

    private readonly struct CsvTextPartition
    {
        internal CsvTextPartition(int start, int end, int rowCount)
        {
            Start = start;
            End = end;
            RowCount = rowCount;
        }

        internal int Start { get; }

        internal int End { get; }

        internal int RowCount { get; }
    }
}
#endif
