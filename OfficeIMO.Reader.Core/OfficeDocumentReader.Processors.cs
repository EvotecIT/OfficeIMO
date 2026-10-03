using System;
using System.Collections.Generic;
using System.Linq;
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
        bool computeHashes,
        CancellationToken cancellationToken) {
        if (ProcessorPipeline.Count == 0) return document;
        OfficeDocumentSource source = SnapshotSource(document);
        ProcessedChunkAggregateSnapshot aggregates = DocumentReaderEngine.CaptureProcessedChunkAggregates(document);
        OfficeDocumentReadResult processed = ProcessorPipeline
            .Process(document, _processingOptions, cancellationToken)
            .Document;
        return DocumentReaderEngine.RefreshProcessedChunks(processed, source, aggregates, computeHashes, cancellationToken);
    }

    private async Task<OfficeDocumentReadResult> ExecuteProcessedDocumentAsync(
        Func<Task<OfficeDocumentReadResult>> read,
        bool computeHashes,
        CancellationToken cancellationToken,
        ReaderOptions? options = null) {
        return await ExecuteAsync(async () => {
            using var readScope = ReaderReadScope.Enter(options);
            OfficeDocumentReadResult document = await read().ConfigureAwait(false);
            if (ProcessorPipeline.Count == 0) return ReaderReadScope.Complete(document);
            OfficeDocumentSource source = SnapshotSource(document);
            ProcessedChunkAggregateSnapshot aggregates = DocumentReaderEngine.CaptureProcessedChunkAggregates(document);
            OfficeDocumentProcessingResult processed = await ProcessorPipeline
                .ProcessAsync(document, _processingOptions, cancellationToken)
                .ConfigureAwait(false);
            return ReaderReadScope.Complete(DocumentReaderEngine.RefreshProcessedChunks(processed.Document, source, aggregates, computeHashes, cancellationToken));
        }, cancellationToken).ConfigureAwait(false);
    }

    private async Task<IReadOnlyList<ReaderChunk>> ExecuteProcessedChunksAsync(
        Func<Task<OfficeDocumentReadResult>> read,
        bool computeHashes,
        CancellationToken cancellationToken,
        ReaderOptions? options = null) {
        OfficeDocumentReadResult document = await ExecuteProcessedDocumentAsync(
            read,
            computeHashes,
            cancellationToken, options).ConfigureAwait(false);
        return document.Chunks ?? Array.Empty<ReaderChunk>();
    }

    private static OfficeDocumentSource SnapshotSource(OfficeDocumentReadResult document) {
        OfficeDocumentSource source = document.Source ?? new OfficeDocumentSource();
        return new OfficeDocumentSource {
            Path = source.Path,
            SourceId = source.SourceId,
            SourceHash = source.SourceHash,
            LastWriteUtc = source.LastWriteUtc,
            LengthBytes = source.LengthBytes,
            Title = source.Title,
            Author = source.Author,
            Subject = source.Subject,
            Keywords = source.Keywords
        };
    }
}
