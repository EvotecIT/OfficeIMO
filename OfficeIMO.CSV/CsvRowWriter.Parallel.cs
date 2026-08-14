#nullable enable

using System.Data;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Data;

namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    /// <summary>
    /// Writes an <see cref="IDataReader"/> by reading it sequentially and formatting
    /// bounded row batches in parallel. Output order always matches source order.
    /// </summary>
    /// <param name="reader">Source reader positioned before the first row.</param>
    /// <param name="parallelOptions">Parallel worker and batch limits.</param>
    /// <param name="cancellationToken">Token observed while reading, formatting, and committing batches.</param>
    /// <remarks>
    /// Data readers are consumed by one thread. Only detached row formatting runs
    /// concurrently, so providers do not need to support concurrent access. The
    /// pipeline retains at most two row batches and clones mutable byte and character
    /// arrays before formatting. Readers that expose custom or otherwise unsafe field
    /// types use the established sequential writer to preserve provider-owned values.
    /// </remarks>
    public void WriteDataReaderParallel(
        IDataReader reader,
        CsvWriteParallelOptions? parallelOptions = null,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        parallelOptions ??= new CsvWriteParallelOptions();
        int degreeOfParallelism = parallelOptions.GetDegreeOfParallelism();
        parallelOptions.GetBatchSize();
        if (degreeOfParallelism == 1)
        {
            WriteDataReader(reader, cancellationToken);
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        int fieldCount = reader.FieldCount;
        if (fieldCount <= 0)
        {
            throw new InvalidOperationException("Data reader must expose at least one field.");
        }

        if (!ParallelRowMappingExtensions.TryCreateIndependentSnapshotPlan(
                reader,
                parallelOptions.MaximumBufferedCellsPerBatch,
                out bool[] cloneColumns,
                out bool fieldLimitExceeded))
        {
            WriteDataReader(reader, cancellationToken);
            return;
        }

        if (fieldLimitExceeded)
        {
            throw parallelOptions.CreateFieldCountLimitException(fieldCount);
        }
        int batchSize = parallelOptions.GetBatchSize(fieldCount);

        var columns = new string[fieldCount];
        for (int index = 0; index < fieldCount; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            columns[index] = reader.GetName(index);
        }

        EnsureColumns(columns);

        object[][] currentRows = CreateParallelRows(batchSize, fieldCount);
        object[][] nextRows = CreateParallelRows(batchSize, fieldCount);

        int workerCapacity = Math.Min(degreeOfParallelism, batchSize);
        var workerBuffers = new StringBuilder[workerCapacity];
        var workerWriters = new StringWriter[workerCapacity];
        for (int workerIndex = 0; workerIndex < workerCapacity; workerIndex++)
        {
            var buffer = new StringBuilder(4096);
            workerBuffers[workerIndex] = buffer;
            workerWriters[workerIndex] = new StringWriter(buffer, _options.Culture);
        }

        try
        {
            bool useBulkValues = true;
            int currentRowCount = ReadParallelBatch(
                reader,
                currentRows,
                fieldCount,
                cloneColumns,
                ref useBulkValues,
                cancellationToken);
            while (currentRowCount != 0)
            {
                Task[] formattingTasks = StartFormattingParallelBatch(
                    currentRows,
                    currentRowCount,
                    workerBuffers,
                    workerWriters,
                    cancellationToken);

                int nextRowCount;
                try
                {
                    nextRowCount = currentRowCount == currentRows.Length
                        ? ReadParallelBatch(
                            reader,
                            nextRows,
                            fieldCount,
                            cloneColumns,
                            ref useBulkValues,
                            cancellationToken,
                            formattingTasks)
                        : 0;
                }
                catch
                {
                    WaitForParallelFormatting(formattingTasks, cancellationToken);
                    throw;
                }

                WaitForParallelFormatting(formattingTasks, cancellationToken);
                CommitParallelBatch(workerBuffers, currentRowCount, cancellationToken);

                object[][] swap = currentRows;
                currentRows = nextRows;
                nextRows = swap;
                currentRowCount = nextRowCount;
            }
        }
        finally
        {
            for (int workerIndex = 0; workerIndex < workerWriters.Length; workerIndex++)
            {
                workerWriters[workerIndex].Dispose();
            }
        }
    }

    private static object[][] CreateParallelRows(int batchSize, int fieldCount)
    {
        var rows = new object[batchSize][];
        for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            rows[rowIndex] = new object[fieldCount];
        }

        return rows;
    }

    private static int ReadParallelBatch(
        IDataReader reader,
        object[][] rows,
        int fieldCount,
        bool[] cloneColumns,
        ref bool useBulkValues,
        CancellationToken cancellationToken,
        Task[]? concurrentFormattingTasks = null)
    {
        int rowCount = 0;
        while (rowCount < rows.Length)
        {
            ThrowIfParallelFormattingFailed(concurrentFormattingTasks, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            if (!reader.Read())
            {
                break;
            }

            ThrowIfParallelFormattingFailed(concurrentFormattingTasks, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            object[] values = rows[rowCount];
            if (useBulkValues && TryGetReaderValues(reader, values))
            {
                CloneMutableParallelValues(values, cloneColumns);
                rowCount++;
                continue;
            }

            useBulkValues = false;
            for (int fieldIndex = 0; fieldIndex < fieldCount; fieldIndex++)
            {
                values[fieldIndex] = reader.GetValue(fieldIndex);
            }

            CloneMutableParallelValues(values, cloneColumns);

            rowCount++;
        }

        return rowCount;
    }

    private static void CloneMutableParallelValues(object[] values, bool[] cloneColumns)
    {
        for (int fieldIndex = 0; fieldIndex < cloneColumns.Length; fieldIndex++)
        {
            if (!cloneColumns[fieldIndex])
            {
                continue;
            }

            if (values[fieldIndex] is byte[] bytes)
            {
                values[fieldIndex] = (byte[])bytes.Clone();
            }
            else if (values[fieldIndex] is char[] characters)
            {
                values[fieldIndex] = (char[])characters.Clone();
            }
        }
    }

    private Task[] StartFormattingParallelBatch(
        object[][] rows,
        int rowCount,
        StringBuilder[] workerBuffers,
        StringWriter[] workerWriters,
        CancellationToken cancellationToken)
    {
        int workerCount = Math.Min(workerBuffers.Length, rowCount);
        var tasks = new Task[workerCount];
        for (int workerIndex = 0; workerIndex < workerCount; workerIndex++)
        {
            int capturedWorkerIndex = workerIndex;
            tasks[workerIndex] = Task.Run(() =>
            {
                int start = (int)((long)rowCount * capturedWorkerIndex / workerCount);
                int end = (int)((long)rowCount * (capturedWorkerIndex + 1) / workerCount);
                StringBuilder buffer = workerBuffers[capturedWorkerIndex];
                buffer.Clear();
                for (int rowIndex = start; rowIndex < end; rowIndex++)
                {
                    if ((rowIndex & 63) == 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    AppendParallelRow(workerWriters[capturedWorkerIndex], buffer, rows[rowIndex]);
                }
            }, cancellationToken);
        }

        return tasks;
    }

    private static void ThrowIfParallelFormattingFailed(
        Task[]? tasks,
        CancellationToken cancellationToken)
    {
        if (tasks == null)
        {
            return;
        }

        for (int taskIndex = 0; taskIndex < tasks.Length; taskIndex++)
        {
            if (tasks[taskIndex].IsFaulted || tasks[taskIndex].IsCanceled)
            {
                WaitForParallelFormatting(tasks, cancellationToken);
                return;
            }
        }
    }

    private static void WaitForParallelFormatting(
        Task[] tasks,
        CancellationToken cancellationToken)
    {
        try
        {
            Task.WaitAll(tasks);
        }
        catch (AggregateException) when (cancellationToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(cancellationToken);
        }
        catch (AggregateException exception) when (exception.InnerExceptions.Count == 1)
        {
            ExceptionDispatchInfo.Capture(exception.InnerExceptions[0]).Throw();
        }
    }

    private void AppendParallelRow(StringWriter writer, StringBuilder buffer, object[] values)
    {
        if (_useDefaultWritePath)
        {
            CsvWriter.AppendRecordDefault(buffer, values, _delimiter, _options.NewLine, _options.Culture);
            return;
        }

        if (_useTextDelimiter)
        {
            CsvWriter.WriteRecord(
                writer,
                values,
                _delimiterText,
                _options.NewLine,
                _options.Culture,
                _options.FormulaInjectionPolicy,
                _options.QuoteMode,
                _quoteFields,
                _columns,
                _options.DateTimeFormat,
                _options.UseUtc,
                _options.NullValue);
            return;
        }

        CsvWriter.WriteRecord(
            writer,
            values,
            _delimiter,
            _options.NewLine,
            _options.Culture,
            _options.FormulaInjectionPolicy,
            _options.QuoteMode,
            _quoteFields,
            _columns,
            _options.DateTimeFormat,
            _options.UseUtc,
            _options.NullValue);
    }

    private void CommitParallelBatch(
        StringBuilder[] workerBuffers,
        int rowCount,
        CancellationToken cancellationToken)
    {
        int workerCount = Math.Min(workerBuffers.Length, rowCount);
        for (int workerIndex = 0; workerIndex < workerCount; workerIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            StringBuilder buffer = workerBuffers[workerIndex];
            if (_stringWriterBuffer != null)
            {
                _stringWriterBuffer.Append(buffer);
                buffer.Clear();
            }
            else
            {
#if NET6_0_OR_GREATER
                CsvWriter.FlushBufferedContent(_writer, buffer);
#else
                _writer.Write(buffer.ToString());
                buffer.Clear();
#endif
            }
        }
    }
}
