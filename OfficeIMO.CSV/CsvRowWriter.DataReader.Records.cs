#if NET8_0_OR_GREATER
using System.Data;
using System.Text;

namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    /// <summary>
    /// Writes typed data-reader rows through a pooled character buffer: common values are formatted straight
    /// into it instead of through a row string builder, and completed rows reach the writer in batches. Only
    /// completed rows are written, so failures and cancellation leave the same output as before.
    /// </summary>
    private void WriteDataReaderRecords(IDataReader reader, CsvWriter.DataReaderFieldKind[] fieldKinds, CancellationToken cancellationToken)
    {
        var scratch = new StringBuilder();
        using var output = new CsvWriter.RecordBuffer();
        var completed = 0;
        bool boxedValues = reader.GetType() == typeof(DataTableReader);
        bool invariantRows = !_useFormattedDefaultWritePath && _delimiter == ','
            && ReferenceEquals(_options.Culture, CultureInfo.InvariantCulture) && CsvWriter.HasOnlyInvariantRecordKinds(fieldKinds);
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
                if (!invariantRows || !(boxedValues
                        ? TryAppendInvariantRecord(output, new CsvWriter.BoxedRecordFields(reader), fieldKinds)
                        : TryAppendInvariantRecord(output, new CsvWriter.TypedRecordFields(reader), fieldKinds)))
                {
                    // After a rejected typed read every later row takes the per-field path, as before.
                    invariantRows = false;
                    CsvWriter.AppendDataReaderRecord(output, scratch, reader, fieldKinds, _delimiter, _options.NewLine,
                        _options.Culture, _useFormattedDefaultWritePath, _options.DateTimeFormat, _options.UseUtc, _options.NullValue, boxedValues);
                }

                completed = output.Length;
                if (completed >= CsvWriter.DataReaderFlushThreshold)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    _writer.Write(output.Written);
                    output.Length = completed = 0;
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            _writer.Write(output.Written);
            output.Length = 0;
        }
        catch
        {
            output.Length = completed;
            if (completed != 0)
            {
                _writer.Write(output.Written);
            }

            throw;
        }
    }

    /// <summary>Writes one row through the one-pass invariant path; false rewinds it for the per-field path.</summary>
    private bool TryAppendInvariantRecord<TFields>(CsvWriter.RecordBuffer output, TFields fields, CsvWriter.DataReaderFieldKind[] fieldKinds)
        where TFields : struct, CsvWriter.IRecordFields
    {
        int rowStart = output.Length;
        try
        {
            CsvWriter.AppendInvariantRecord(output, ref fields, fieldKinds, _options.NewLine);
            return true;
        }
        catch (Exception exception) when (exception is InvalidCastException or NotSupportedException or NotImplementedException)
        {
            output.Length = rowStart;
            return false;
        }
    }
}
#endif
