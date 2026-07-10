using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Processes an already-read document through this reader's frozen pipeline.</summary>
    public OfficeDocumentProcessingResult ProcessDocument(
        OfficeDocumentReadResult document,
        CancellationToken cancellationToken = default) {
        return ProcessorPipeline.Process(document, _processingOptions, cancellationToken);
    }

    /// <summary>Processes an already-read document asynchronously through this reader's frozen pipeline.</summary>
    public Task<OfficeDocumentProcessingResult> ProcessDocumentAsync(
        OfficeDocumentReadResult document,
        CancellationToken cancellationToken = default) {
        return ProcessorPipeline.ProcessAsync(document, _processingOptions, cancellationToken);
    }

    private OfficeDocumentReadResult ProcessDocumentResult(
        OfficeDocumentReadResult document,
        CancellationToken cancellationToken) {
        if (ProcessorPipeline.Count == 0) return document;
        return ProcessorPipeline.Process(document, _processingOptions, cancellationToken).Document;
    }

    private async Task<OfficeDocumentReadResult> ExecuteProcessedDocumentAsync(
        Func<Task<OfficeDocumentReadResult>> read,
        CancellationToken cancellationToken) {
        return await ExecuteAsync(async () => {
            OfficeDocumentReadResult document = await read().ConfigureAwait(false);
            if (ProcessorPipeline.Count == 0) return document;
            OfficeDocumentProcessingResult processed = await ProcessorPipeline
                .ProcessAsync(document, _processingOptions, cancellationToken)
                .ConfigureAwait(false);
            return processed.Document;
        }, cancellationToken).ConfigureAwait(false);
    }

    private async Task<IReadOnlyList<ReaderChunk>> ExecuteProcessedChunksAsync(
        Func<Task<OfficeDocumentReadResult>> read,
        CancellationToken cancellationToken) {
        OfficeDocumentReadResult document = await ExecuteProcessedDocumentAsync(read, cancellationToken).ConfigureAwait(false);
        return document.Chunks ?? Array.Empty<ReaderChunk>();
    }
}
