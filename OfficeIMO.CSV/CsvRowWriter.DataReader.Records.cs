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
                CsvWriter.AppendDataReaderRecord(output, scratch, reader, fieldKinds, _delimiter, _options.NewLine,
                    _options.Culture, _useFormattedDefaultWritePath, _options.DateTimeFormat, _options.UseUtc, _options.NullValue);
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
}
#endif
