using System.Buffers;
using System.Data;

namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    /// <summary>
    /// Writes all rows from an <see cref="IDataReader"/> using the reader field names as CSV columns.
    /// </summary>
    /// <param name="reader">Source data reader positioned before the first row.</param>
    /// <remarks>
    /// The method streams rows without materializing a document and reuses one row buffer for the whole reader.
    /// </remarks>
    public void WriteDataReader(IDataReader reader) =>
        WriteDataReader(reader, CancellationToken.None);

    /// <summary>
    /// Writes the rows of a <see cref="DataTable"/> using its column names as CSV columns: the rows in order without
    /// the deleted ones, each value from the row's current version, as writing the table's data reader does.
    /// </summary>
    /// <param name="table">Source table with at least one column.</param>
    public void WriteDataTable(DataTable table)
    {
        ArgumentNullException.ThrowIfNull(table);
        using var reader = new CsvDataTableRowReader(table);
        WriteDataReader(reader, CancellationToken.None);
    }

    /// <summary>
    /// Writes all rows from an <see cref="IDataReader"/> using the reader field names as CSV columns.
    /// </summary>
    /// <param name="reader">Source data reader positioned before the first row.</param>
    /// <param name="cancellationToken">Token observed while projecting and writing rows.</param>
    /// <remarks>
    /// The method streams rows without materializing a document and reuses one row buffer for the whole reader.
    /// </remarks>
    public void WriteDataReader(IDataReader reader, CancellationToken cancellationToken)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(reader);

        cancellationToken.ThrowIfCancellationRequested();
        var fieldCount = reader.FieldCount;
        if (fieldCount <= 0)
        {
            throw new InvalidOperationException("Data reader must expose at least one field.");
        }

        var columns = new string[fieldCount];
        for (var i = 0; i < fieldCount; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            columns[i] = reader.GetName(i);
        }

        EnsureColumns(columns);

        // Typed values keep the default quoting path, with or without date and null formatting.
        if ((_useDefaultWritePath || _useFormattedDefaultWritePath)
            && CsvWriter.TryCreateDataReaderFieldKinds(reader) is { } fieldKinds)
        {
            WriteDataReaderRecords(reader, fieldKinds, cancellationToken);
            return;
        }

        var usesBatchedRecords = _batchFormattedTextDataReader || _useDefaultWritePath;
        int flushThreshold = _useDefaultWritePath
            ? CsvWriter.DataReaderFlushThreshold
            : TextDelimiterDataReaderFlushThreshold;
        if (usesBatchedRecords)
        {
            _rowBuffer.Clear();
        }

        var rowValues = new object[fieldCount];
        var useBufferedValues = true;
        char[]? batch = null;
        var batchLength = 0;
        void FlushBatch()
        {
            if (batchLength == 0) return;
            int length = batchLength;
            batchLength = 0;
            _writer.Write(batch!, 0, length);
        }
        void BufferCompletedRow()
        {
            batch ??= ArrayPool<char>.Shared.Rent(flushThreshold);
            int rowLength = _rowBuffer.Length;
            if (rowLength > batch.Length - batchLength) FlushBatch();
            if (rowLength > batch.Length)
            {
                CsvWriter.FlushBufferedContent(_writer, _rowBuffer);
                cancellationToken.ThrowIfCancellationRequested();
            }
            else
            {
                _rowBuffer.CopyTo(0, batch, batchLength, rowLength);
                batchLength += rowLength;
                _rowBuffer.Clear();
                if (batchLength >= flushThreshold)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    FlushBatch();
                }
            }
        }
        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!reader.Read())
                {
                    break;
                }
                cancellationToken.ThrowIfCancellationRequested();
                bool haveValues = useBufferedValues && TryGetReaderValues(reader, rowValues);
                if (!haveValues)
                {
                    useBufferedValues = false;
                    if (usesBatchedRecords)
                    {
                        for (int index = 0; index < fieldCount; index++) rowValues[index] = reader.GetValue(index);
                    }
                }
                if (usesBatchedRecords)
                {
                    CsvWriter.AppendDataReaderRecordBuffered(
                        _rowBuffer, rowValues, _delimiterText, _options.NewLine, _options.Culture,
                        _options.FormulaInjectionPolicy, _options.ResolvedQuoteMode, _quoteFields, _columns,
                        _options.DateTimeFormat, _options.UseUtc, _options.NullValue);
                    BufferCompletedRow();
                }
                else if (haveValues)
                {
                    WriteBuffered(rowValues);
                }
                else
                {
                    WriteBuffered(fieldCount, reader, static (record, index) =>
                    {
                        var value = record.GetValue(index);
                        return ReferenceEquals(value, DBNull.Value) ? null : value;
                    });
                }
            }

            if (usesBatchedRecords)
            {
                cancellationToken.ThrowIfCancellationRequested();
                FlushBatch();
            }
        }
        catch
        {
            if (usesBatchedRecords)
            {
                // Drop the unfinished row and deliver the completed ones.
                _rowBuffer.Clear();
                FlushBatch();
            }

            throw;
        }
        finally
        {
            if (batch != null)
                ArrayPool<char>.Shared.Return(batch, clearArray: true);
        }
    }

}
