using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

public static partial class DocumentReader {
    internal static OfficeDocumentReadResult RefreshProcessedChunks(
        OfficeDocumentReadResult document,
        OfficeDocumentSource sourceFallback,
        bool computeHashes,
        CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceFallback == null) throw new ArgumentNullException(nameof(sourceFallback));

        IReadOnlyList<ReaderChunk> chunks = document.Chunks ?? Array.Empty<ReaderChunk>();
        OfficeDocumentSource documentSource = document.Source ?? new OfficeDocumentSource();
        ReaderChunk? firstChunk = chunks.FirstOrDefault(chunk => chunk != null);
        string sourcePath = FirstNonEmpty(
            documentSource.Path,
            sourceFallback.Path,
            firstChunk?.Location?.Path,
            "memory");
        string sourceId = FirstNonEmpty(
            documentSource.SourceId,
            sourceFallback.SourceId,
            null,
            BuildSourceId(sourcePath));
        string? sourceHash = FirstNonEmptyOrNull(
            documentSource.SourceHash,
            sourceFallback.SourceHash,
            null);
        DateTime? sourceLastWriteUtc = documentSource.LastWriteUtc ?? sourceFallback.LastWriteUtc;
        long? sourceLengthBytes = documentSource.LengthBytes ?? sourceFallback.LengthBytes;

        document.Source = documentSource;
        if (string.IsNullOrWhiteSpace(documentSource.Path)) documentSource.Path = sourcePath;
        if (string.IsNullOrWhiteSpace(documentSource.SourceId)) documentSource.SourceId = sourceId;
        if (string.IsNullOrWhiteSpace(documentSource.SourceHash)) documentSource.SourceHash = sourceHash;
        documentSource.LastWriteUtc ??= sourceLastWriteUtc;
        documentSource.LengthBytes ??= sourceLengthBytes;
        if (string.IsNullOrWhiteSpace(documentSource.Title)) documentSource.Title = sourceFallback.Title;
        if (string.IsNullOrWhiteSpace(documentSource.Author)) documentSource.Author = sourceFallback.Author;
        if (string.IsNullOrWhiteSpace(documentSource.Subject)) documentSource.Subject = sourceFallback.Subject;
        if (string.IsNullOrWhiteSpace(documentSource.Keywords)) documentSource.Keywords = sourceFallback.Keywords;

        for (int index = 0; index < chunks.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            ReaderChunk chunk = chunks[index]
                ?? throw new InvalidOperationException($"Processed document chunk at index {index} is null.");
            chunk.SourceId = sourceId;
            chunk.SourceHash = sourceHash;
            chunk.SourceLastWriteUtc = sourceLastWriteUtc;
            chunk.SourceLengthBytes = sourceLengthBytes;
            chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
            chunk.ChunkHash = computeHashes ? ComputeChunkHash(chunk) : null;
        }

        document.Chunks = chunks;
        return document;
    }

    private static string FirstNonEmpty(string? first, string? second, string? third, string fallback) {
        if (!string.IsNullOrWhiteSpace(first)) return first!;
        if (!string.IsNullOrWhiteSpace(second)) return second!;
        if (!string.IsNullOrWhiteSpace(third)) return third!;
        return fallback;
    }

    private static string? FirstNonEmptyOrNull(string? first, string? second, string? third) {
        if (!string.IsNullOrWhiteSpace(first)) return first;
        if (!string.IsNullOrWhiteSpace(second)) return second;
        return string.IsNullOrWhiteSpace(third) ? null : third;
    }
}
