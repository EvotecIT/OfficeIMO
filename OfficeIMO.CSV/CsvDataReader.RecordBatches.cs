#nullable enable
#if NET8_0_OR_GREATER
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

internal sealed partial class CsvDataReader
{
    internal async ValueTask<(CsvRecordTextBatch Batch, bool ReachedEnd)> ReadRecordTextBatchAsync(
        int preferredRows, CancellationToken cancellationToken)
    {
        var batch = new CsvRecordTextBatch(preferredRows, FieldCount, _culture, _dateTimeFormats);
        try
        {
            bool reachedEnd = false;
            while (batch.Count < batch.RowCapacity)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!await ReadAsync(cancellationToken).ConfigureAwait(false))
                {
                    reachedEnd = true;
                    break;
                }
                cancellationToken.ThrowIfCancellationRequested();
                batch.CaptureRow(this);
            }
            cancellationToken.ThrowIfCancellationRequested();
            return (batch, reachedEnd);
        }
        catch { batch.Dispose(); throw; }
    }
}
#endif
