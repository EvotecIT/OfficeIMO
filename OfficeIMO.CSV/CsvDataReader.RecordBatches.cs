#if NET8_0_OR_GREATER
using System.Text;

namespace OfficeIMO.CSV;

public partial class CsvDataReader
{
    /// <summary>Advances an internal record consumer without adapting completed incremental reads to Task.</summary>
    internal ValueTask<bool> ReadRecordAsync(CancellationToken cancellationToken) =>
        _asyncTextRowSource is { } asynchronousRows
            ? ReadIncrementalAsync(asynchronousRows, asynchronous: true, cancellationToken)
            : new ValueTask<bool>(ReadAsync(cancellationToken));

    /// <summary>Creates bounded storage whose caller owns its complete worker lifetime and disposal.</summary>
    internal CsvRecordTextBatch CreateRecordTextBatch(int preferredRows, bool captureUtf8 = false, bool includePositions = false)
    {
        Encoding? encoding = captureUtf8 ? _stringRowOptions?.Encoding ?? Encoding.UTF8 : null;
        // A custom encoding can have caller-owned behavior; keep its decoding on the source.
        if (encoding is not null && !ReferenceEquals(encoding, Encoding.UTF8) &&
            encoding.GetType() != typeof(UTF8Encoding)) encoding = null;
        return new(preferredRows, FieldCount, _culture, _dateTimeFormats, encoding, includePositions);
    }

    /// <summary>Resets and fills caller-owned storage after its previous worker has completed.</summary>
    /// <returns>Whether the source ended while filling the batch.</returns>
    [System.Runtime.CompilerServices.AsyncMethodBuilder(typeof(System.Runtime.CompilerServices.PoolingAsyncValueTaskMethodBuilder<>))]
    internal async ValueTask<bool> ReadRecordTextBatchAsync(
        CsvRecordTextBatch batch, CancellationToken cancellationToken, bool asynchronous = true)
    {
        batch.Reset();
        bool reachedEnd = false;
        while (batch.Count < batch.RowCapacity)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!batch.CapturesPositions && _useRawStringValues && !_hasBufferedRow &&
                _textRowSource is CsvIncrementalRowSource rawSource && rawSource.CanCaptureRawRecords)
            {
                var raw = await ReadRawRecordRunAsync(rawSource, batch, cancellationToken).ConfigureAwait(false);
                if (raw.ReachedEnd) { reachedEnd = true; break; }
                if (raw.Count > 0) continue;
                // The same batch captures complex records and text-cap overflow canonically.
                // Never dispatch a raw prefix before a later row has parsed successfully.
            }
            bool available = asynchronous
                ? await ReadRecordAsync(cancellationToken).ConfigureAwait(false)
                : ReadCore(cancellationToken);
            if (!available)
            {
                reachedEnd = true;
                break;
            }
            cancellationToken.ThrowIfCancellationRequested();
            // Converted/static values need the reader's canonical null semantics.
            // Ordinary incremental records already expose aligned source text and metadata.
            if (_useRawStringValues && _textRowSource is CsvIncrementalRowSource incremental)
                batch.CaptureRow(incremental, _stringNullValue);
            else if (batch.CanCaptureUtf8 && _useRawStringValues &&
                _textRowSource is CsvParser.CsvUtf8StreamDataReaderRowSource utf8Rows)
                batch.CaptureRow(utf8Rows, _stringNullValue);
            else
                batch.CaptureRow(this);
            if (batch.CapturesPositions) batch.CapturePosition(PhysicalLineNumber, PhysicalEndLineNumber);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return reachedEnd;
    }

    [System.Runtime.CompilerServices.AsyncMethodBuilder(typeof(System.Runtime.CompilerServices.PoolingAsyncValueTaskMethodBuilder<>))]
    private async ValueTask<(int Count, bool ReachedEnd)> ReadRawRecordRunAsync(
        CsvIncrementalRowSource source, CsvRecordTextBatch batch, CancellationToken cancellationToken)
    {
        if (_closed) return (0, true);
        if (_readFailed) throw new InvalidOperationException("The CSV reader cannot continue after a failed advance.");
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            _processingCancellationToken.ThrowIfCancellationRequested();
            ClearCurrentRow();
            var capture = await source.CaptureRawRecordRunAsync(batch, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            _processingCancellationToken.ThrowIfCancellationRequested();
            SetHasCurrentTextRow(capture.Count > 0);
            if (capture.Count > 0) { _hasRows ??= true; _rowIndex += capture.Count; }
            else if (capture.ReachedEnd) _hasRows ??= false;
            return capture;
        }
        catch { MarkReadFailed(); throw; }
    }
}
#endif
