using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>Per-operation input accounting around the existing native read admission policy.</summary>
internal sealed class ReaderBatchReadPolicy {
    private readonly ReaderOptions? _options;
    private long? _remaining;
    internal ReaderBatchOptions? BatchOptions { get; }

    internal ReaderBatchReadPolicy(ReaderOptions? options, ReaderBatchOptions? batch) {
        _options = options;
        if (batch?.MaxTotalBytes is long maximum) {
            if (batch.MaxDegreeOfParallelism < 1) throw new ArgumentOutOfRangeException(nameof(batch.MaxDegreeOfParallelism));
            if (maximum < 1) throw new ArgumentOutOfRangeException(nameof(batch.MaxTotalBytes));
            _remaining = maximum;
            // A remaining byte allowance depends on the preceding read. Keep the same outer read scope and its other budgets.
            BatchOptions = new ReaderBatchOptions { MaxDocuments = batch.MaxDocuments, MaxDegreeOfParallelism = 1, MaxTotalBytes = maximum };
        } else BatchOptions = batch;
    }

    internal async Task<OfficeDocumentReadResult> ReadAsync(string path,
        Func<string, ReaderOptions?, CancellationToken, Task<OfficeDocumentReadResult>> read, CancellationToken token) {
        ReaderOptions? options = _options;
        if (_remaining.HasValue) {
            token.ThrowIfCancellationRequested();
            if (_remaining.Value == 0) throw new IOException("Batch MaxTotalBytes is exhausted; the source was not opened.");
            options = DocumentReaderEngine.ApplyFolderInputBudget(_options, _remaining.Value);
        }
        OfficeDocumentReadResult document = await read(path, options, token).ConfigureAwait(false);
        if (_remaining.HasValue) {
            long? length = document.Source?.LengthBytes;
            if (!length.HasValue || length.Value < 0 || length.Value > _remaining.Value)
                throw new InvalidDataException("A byte-bounded batch requires the reader to report a source length within its remaining input allowance.");
            _remaining -= length.Value;
        }
        return document;
    }
}
