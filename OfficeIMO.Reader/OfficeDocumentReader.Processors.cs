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
        CancellationToken cancellationToken) {
        return await ExecuteAsync(async () => {
            OfficeDocumentReadResult document = await read().ConfigureAwait(false);
            if (ProcessorPipeline.Count == 0) return document;
            OfficeDocumentSource source = SnapshotSource(document);
            ProcessedChunkAggregateSnapshot aggregates = DocumentReaderEngine.CaptureProcessedChunkAggregates(document);
            OfficeDocumentProcessingResult processed = await ProcessorPipeline
                .ProcessAsync(document, _processingOptions, cancellationToken)
                .ConfigureAwait(false);
            return DocumentReaderEngine.RefreshProcessedChunks(processed.Document, source, aggregates, computeHashes, cancellationToken);
        }, cancellationToken).ConfigureAwait(false);
    }

    private async Task<IReadOnlyList<ReaderChunk>> ExecuteProcessedChunksAsync(
        Func<Task<OfficeDocumentReadResult>> read,
        bool computeHashes,
        CancellationToken cancellationToken) {
        OfficeDocumentReadResult document = await ExecuteProcessedDocumentAsync(
            read,
            computeHashes,
            cancellationToken).ConfigureAwait(false);
        return document.Chunks ?? Array.Empty<ReaderChunk>();
    }

    private static OfficeDocumentSource SnapshotSource(OfficeDocumentReadResult document) {
        OfficeDocumentSource source = document.Source ?? new OfficeDocumentSource();
        ReaderChunk? firstChunk = document.Chunks?.FirstOrDefault(chunk => chunk != null);
        return new OfficeDocumentSource {
            Path = string.IsNullOrWhiteSpace(source.Path) ? firstChunk?.Location?.Path : source.Path,
            SourceId = string.IsNullOrWhiteSpace(source.SourceId) ? firstChunk?.SourceId : source.SourceId,
            SourceHash = string.IsNullOrWhiteSpace(source.SourceHash) ? firstChunk?.SourceHash : source.SourceHash,
            LastWriteUtc = source.LastWriteUtc ?? firstChunk?.SourceLastWriteUtc,
            LengthBytes = source.LengthBytes ?? firstChunk?.SourceLengthBytes,
            Title = source.Title,
            Author = source.Author,
            Subject = source.Subject,
            Keywords = source.Keywords
        };
    }
}
