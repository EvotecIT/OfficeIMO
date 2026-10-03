using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>Processes an already-read document and nested results through this reader's frozen pipeline.</summary>
    public OfficeDocumentProcessingResult ProcessDocument(
        OfficeDocumentReadResult document,
        CancellationToken cancellationToken = default) {
        OfficeDocumentProcessingResult result = ProcessorPipeline.Process(document, _processingOptions, cancellationToken);
        if (ProcessorPipeline.Count != 0) {
            ReaderReadScope.AttachPendingNested(result.Document);
            var ancestors = new HashSet<OfficeDocumentReadResult> { document, result.Document };
            OfficeDocumentProcessorStepResult[] steps = result.Steps.ToArray();
            ProcessNestedDocuments(result.Document, false, cancellationToken, ancestors, 0, steps);
            return new OfficeDocumentProcessingResult(result.Document, steps);
        }
        return result;
    }

    /// <summary>Processes an already-read document and nested results asynchronously through this reader's frozen pipeline.</summary>
    public async Task<OfficeDocumentProcessingResult> ProcessDocumentAsync(
        OfficeDocumentReadResult document,
        CancellationToken cancellationToken = default) {
        OfficeDocumentProcessingResult result = await ProcessorPipeline
            .ProcessAsync(document, _processingOptions, cancellationToken).ConfigureAwait(false);
        if (ProcessorPipeline.Count != 0) {
            ReaderReadScope.AttachPendingNested(result.Document);
            var ancestors = new HashSet<OfficeDocumentReadResult> { document, result.Document };
            OfficeDocumentProcessorStepResult[] steps = result.Steps.ToArray();
            await ProcessNestedDocumentsAsync(result.Document, false, cancellationToken, ancestors, 0, steps)
                .ConfigureAwait(false);
            return new OfficeDocumentProcessingResult(result.Document, steps);
        }
        return result;
    }

    private OfficeDocumentReadResult ProcessDocumentResult(
        OfficeDocumentReadResult document,
        bool computeHashes,
        CancellationToken cancellationToken) =>
        ProcessorPipeline.Count == 0 ? document : ProcessDocumentTree(document, computeHashes, cancellationToken,
            new HashSet<OfficeDocumentReadResult>(), 0);

    private OfficeDocumentReadResult ProcessDocumentTree(
        OfficeDocumentReadResult document,
        bool computeHashes,
        CancellationToken cancellationToken,
        HashSet<OfficeDocumentReadResult> ancestors,
        int depth,
        OfficeDocumentProcessorStepResult[]? aggregateSteps = null) {
        if (depth > 64 || !ancestors.Add(document))
            throw new InvalidOperationException("Nested document results cannot contain a cycle or exceed depth 64.");
        try {
            OfficeDocumentSource source = SnapshotSource(document);
            ProcessedChunkAggregateSnapshot aggregates = DocumentReaderEngine.CaptureProcessedChunkAggregates(document);
            OfficeDocumentProcessingResult operation = ProcessorPipeline.Process(document, _processingOptions, cancellationToken);
            if (aggregateSteps != null) MergeSteps(aggregateSteps, operation.Steps);
            OfficeDocumentReadResult processed = DocumentReaderEngine.RefreshProcessedChunks(
                operation.Document, source, aggregates, computeHashes, cancellationToken);
            ReaderReadScope.AttachPendingNested(processed);
            bool replaced = !ReferenceEquals(processed, document);
            if (replaced && !ancestors.Add(processed))
                throw new InvalidOperationException("Nested document results cannot contain a cycle.");
            try {
                ProcessNestedDocuments(processed, computeHashes, cancellationToken, ancestors, depth, aggregateSteps);
                return processed;
            } finally {
                if (replaced) ancestors.Remove(processed);
            }
        } finally { ancestors.Remove(document); }
    }

    private void ProcessNestedDocuments(OfficeDocumentReadResult document, bool computeHashes,
        CancellationToken cancellationToken, HashSet<OfficeDocumentReadResult> ancestors, int depth,
        OfficeDocumentProcessorStepResult[]? aggregateSteps = null) {
        foreach (OfficeDocumentNestedResult nested in document.NestedDocuments ?? Array.Empty<OfficeDocumentNestedResult>()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (nested == null || nested.Document == null)
                throw new InvalidOperationException("A nested document requires a document.");
            nested.Document = ProcessDocumentTree(nested.Document, computeHashes, cancellationToken, ancestors, depth + 1, aggregateSteps);
        }
    }

    private static void MergeSteps(OfficeDocumentProcessorStepResult[] aggregate,
        IReadOnlyList<OfficeDocumentProcessorStepResult> child) {
        for (int index = 0; index < aggregate.Length; index++) {
            if (child[index].Status == OfficeDocumentProcessorStepStatus.Failed ||
                aggregate[index].Status == OfficeDocumentProcessorStepStatus.Completed &&
                child[index].Status == OfficeDocumentProcessorStepStatus.Skipped)
                aggregate[index] = child[index];
        }
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
            OfficeDocumentReadResult processed = await ProcessDocumentTreeAsync(document, computeHashes,
                cancellationToken, new HashSet<OfficeDocumentReadResult>(), 0).ConfigureAwait(false);
            return ReaderReadScope.Complete(processed);
        }, cancellationToken).ConfigureAwait(false);
    }

    private async Task<OfficeDocumentReadResult> ProcessDocumentTreeAsync(
        OfficeDocumentReadResult document,
        bool computeHashes,
        CancellationToken cancellationToken,
        HashSet<OfficeDocumentReadResult> ancestors,
        int depth,
        OfficeDocumentProcessorStepResult[]? aggregateSteps = null) {
        if (depth > 64 || !ancestors.Add(document))
            throw new InvalidOperationException("Nested document results cannot contain a cycle or exceed depth 64.");
        try {
            OfficeDocumentSource source = SnapshotSource(document);
            ProcessedChunkAggregateSnapshot aggregates = DocumentReaderEngine.CaptureProcessedChunkAggregates(document);
            OfficeDocumentProcessingResult operation = await ProcessorPipeline
                .ProcessAsync(document, _processingOptions, cancellationToken).ConfigureAwait(false);
            if (aggregateSteps != null) MergeSteps(aggregateSteps, operation.Steps);
            OfficeDocumentReadResult processed = DocumentReaderEngine.RefreshProcessedChunks(
                operation.Document, source, aggregates, computeHashes, cancellationToken);
            ReaderReadScope.AttachPendingNested(processed);
            bool replaced = !ReferenceEquals(processed, document);
            if (replaced && !ancestors.Add(processed))
                throw new InvalidOperationException("Nested document results cannot contain a cycle.");
            try {
                await ProcessNestedDocumentsAsync(processed, computeHashes, cancellationToken, ancestors, depth, aggregateSteps)
                    .ConfigureAwait(false);
                return processed;
            } finally {
                if (replaced) ancestors.Remove(processed);
            }
        } finally { ancestors.Remove(document); }
    }

    private async Task ProcessNestedDocumentsAsync(OfficeDocumentReadResult document, bool computeHashes,
        CancellationToken cancellationToken, HashSet<OfficeDocumentReadResult> ancestors, int depth,
        OfficeDocumentProcessorStepResult[]? aggregateSteps = null) {
        foreach (OfficeDocumentNestedResult nested in document.NestedDocuments ?? Array.Empty<OfficeDocumentNestedResult>()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (nested == null || nested.Document == null)
                throw new InvalidOperationException("A nested document requires a document.");
            nested.Document = await ProcessDocumentTreeAsync(nested.Document, computeHashes,
                cancellationToken, ancestors, depth + 1, aggregateSteps).ConfigureAwait(false);
        }
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
