using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    internal static OfficeDocumentReadResult RefreshProcessedChunks(
        OfficeDocumentReadResult document,
        OfficeDocumentSource sourceFallback,
        ProcessedChunkAggregateSnapshot aggregateSnapshot,
        bool computeHashes,
        CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (sourceFallback == null) throw new ArgumentNullException(nameof(sourceFallback));
        if (aggregateSnapshot == null) throw new ArgumentNullException(nameof(aggregateSnapshot));

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

        AlignProcessedChunkRepresentations(chunks, aggregateSnapshot);
        for (int index = 0; index < chunks.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            ReaderChunk chunk = chunks[index]
                ?? throw new InvalidOperationException($"Processed document chunk at index {index} is null.");
            chunk.SourceId = sourceId;
            chunk.SourceHash = sourceHash;
            chunk.SourceLastWriteUtc = sourceLastWriteUtc;
            chunk.SourceLengthBytes = sourceLengthBytes;
            chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
            if (computeHashes) chunk.ChunkHash = ComputeChunkHash(chunk);
        }

        document.Chunks = chunks;
        RefreshChunkDerivedAggregates(document, chunks, aggregateSnapshot, cancellationToken);
        return document;
    }

    private static void AlignProcessedChunkRepresentations(
        IReadOnlyList<ReaderChunk> chunks,
        ProcessedChunkAggregateSnapshot snapshot) {
        if (!snapshot.ChunkDerived) return;
        for (int index = 0; index < chunks.Count; index++) {
            ReaderChunk chunk = chunks[index];
            ProcessedChunkState? prior = FindOriginalChunk(chunk, index, snapshot.Chunks);
            if (!prior.HasValue) continue;
            if (prior.Value.ContinuesPreviousChunk) chunk.ContinuesPreviousChunk = true;
            bool textChanged = !string.Equals(chunk.Text, prior.Value.Text, StringComparison.Ordinal);
            bool markdownChanged = !string.Equals(chunk.Markdown, prior.Value.Markdown, StringComparison.Ordinal);
            if (textChanged && !markdownChanged) chunk.Markdown = null;
        }
    }

    internal static ProcessedChunkAggregateSnapshot CaptureProcessedChunkAggregates(OfficeDocumentReadResult document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        IReadOnlyList<ReaderChunk> chunks = document.Chunks ?? Array.Empty<ReaderChunk>();
        IReadOnlyList<OfficeDocumentBlock> blocks = document.Blocks ?? Array.Empty<OfficeDocumentBlock>();
        bool chunkDerived = chunks.Count == blocks.Count;
        if (chunkDerived) {
            for (int index = 0; index < chunks.Count; index++) {
                ReaderChunk chunk = chunks[index];
                OfficeDocumentBlock block = blocks[index];
                string expectedId = !string.IsNullOrWhiteSpace(chunk.Id)
                    ? chunk.Id
                    : "chunk-" + index.ToString("D4", System.Globalization.CultureInfo.InvariantCulture);
                string expectedKind = string.IsNullOrWhiteSpace(chunk.Location.SourceBlockKind)
                    ? "chunk"
                    : chunk.Location.SourceBlockKind!;
                if (block == null ||
                    !string.Equals(block.Id, expectedId, StringComparison.Ordinal) ||
                    !string.Equals(block.Kind, expectedKind, StringComparison.Ordinal) ||
                    !string.Equals(block.Text, chunk.Text ?? string.Empty, StringComparison.Ordinal) ||
                    !ReferenceEquals(block.Location, chunk.Location)) {
                    chunkDerived = false;
                    break;
                }
            }
        }
        if (chunkDerived && !string.Equals(document.Markdown, BuildChunkDocumentMarkdown(chunks), StringComparison.Ordinal)) {
            chunkDerived = false;
        }

        return new ProcessedChunkAggregateSnapshot(
            document,
            chunkDerived,
            document.Markdown,
            chunks.Select(static chunk => new ProcessedChunkState(
                chunk,
                chunk?.Id,
                chunk?.Text,
                chunk?.Markdown,
                chunk?.ContinuesPreviousChunk == true)).ToArray(),
            blocks.Select(static block => new ProcessedBlockState(block, block?.Id, block?.Kind, block?.Text, block?.Location)).ToArray());
    }

    private static void RefreshChunkDerivedAggregates(
        OfficeDocumentReadResult document,
        IReadOnlyList<ReaderChunk> chunks,
        ProcessedChunkAggregateSnapshot snapshot,
        CancellationToken cancellationToken) {
        if (!snapshot.ChunkDerived || !ChunksChanged(chunks, snapshot.Chunks)) return;
        bool sameDocument = ReferenceEquals(document, snapshot.Document);
        if ((sameDocument && string.Equals(document.Markdown, snapshot.Markdown, StringComparison.Ordinal)) ||
            (!sameDocument && string.IsNullOrWhiteSpace(document.Markdown))) {
            document.Markdown = BuildProcessedChunkMarkdown(chunks, snapshot.Chunks);
        }

        bool blocksUnchanged = sameDocument && BlocksRemainOriginal(document.Blocks, snapshot.Blocks);
        if (!blocksUnchanged && (sameDocument || (document.Blocks?.Count ?? 0) > 0)) return;

        OfficeDocumentBlock[] rebuilt = BuildChunkDocumentBlocks(chunks).ToArray();
        bool rebuildPageBlocks = sameDocument && PageBlocksRemainOriginal(document.Pages, snapshot.Blocks);
        document.Blocks = rebuilt;
        if (rebuildPageBlocks) RebuildPageBlocks(document.Pages, rebuilt, cancellationToken);
    }

    private static bool ChunksChanged(
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<ProcessedChunkState> original) {
        if (chunks.Count != original.Count) return true;
        for (int index = 0; index < chunks.Count; index++) {
            ReaderChunk chunk = chunks[index];
            ProcessedChunkState state = original[index];
            if (!ReferenceEquals(chunk, state.Chunk) ||
                !string.Equals(chunk.Id, state.Id, StringComparison.Ordinal) ||
                !string.Equals(chunk.Text, state.Text, StringComparison.Ordinal) ||
                !string.Equals(chunk.Markdown, state.Markdown, StringComparison.Ordinal) ||
                chunk.ContinuesPreviousChunk != state.ContinuesPreviousChunk) return true;
        }
        return false;
    }

    private static bool BlocksRemainOriginal(
        IReadOnlyList<OfficeDocumentBlock>? blocks,
        IReadOnlyList<ProcessedBlockState> original) {
        if (blocks == null || blocks.Count != original.Count) return false;
        for (int index = 0; index < blocks.Count; index++) {
            OfficeDocumentBlock block = blocks[index];
            ProcessedBlockState state = original[index];
            if (!ReferenceEquals(block, state.Block) ||
                !string.Equals(block.Id, state.Id, StringComparison.Ordinal) ||
                !string.Equals(block.Kind, state.Kind, StringComparison.Ordinal) ||
                !string.Equals(block.Text, state.Text, StringComparison.Ordinal) ||
                !ReferenceEquals(block.Location, state.Location)) return false;
        }
        return true;
    }

    private static bool PageBlocksRemainOriginal(
        IReadOnlyList<OfficeDocumentPage>? pages,
        IReadOnlyList<ProcessedBlockState> original) {
        if (pages == null) return false;
        var originalBlocks = new HashSet<OfficeDocumentBlock>(
            original.Where(static state => state.Block != null).Select(static state => state.Block!),
            ReferenceIdentityComparer<OfficeDocumentBlock>.Instance);
        foreach (OfficeDocumentPage page in pages) {
            if (page?.Blocks == null) continue;
            foreach (OfficeDocumentBlock block in page.Blocks) {
                if (block != null && !originalBlocks.Contains(block)) return false;
            }
        }
        return true;
    }

    private static string? BuildProcessedChunkMarkdown(
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<ProcessedChunkState> original) {
        return JoinChunkMarkdown(chunks, (chunk, index) => {
            ProcessedChunkState? prior = FindOriginalChunk(chunk, index, original);
            bool textChanged = prior.HasValue && !string.Equals(chunk.Text, prior.Value.Text, StringComparison.Ordinal);
            bool markdownChanged = prior.HasValue && !string.Equals(chunk.Markdown, prior.Value.Markdown, StringComparison.Ordinal);
            return textChanged && !markdownChanged
                ? chunk.Text
                : (string.IsNullOrWhiteSpace(chunk.Markdown) ? chunk.Text : chunk.Markdown);
        });
    }

    private static ProcessedChunkState? FindOriginalChunk(
        ReaderChunk chunk,
        int index,
        IReadOnlyList<ProcessedChunkState> original) {
        if (index < original.Count && ReferenceEquals(chunk, original[index].Chunk)) return original[index];
        for (int stateIndex = 0; stateIndex < original.Count; stateIndex++) {
            if (ReferenceEquals(chunk, original[stateIndex].Chunk)) return original[stateIndex];
        }
        if (!string.IsNullOrWhiteSpace(chunk.Id)) {
            ProcessedChunkState? match = null;
            for (int stateIndex = 0; stateIndex < original.Count; stateIndex++) {
                if (!string.Equals(chunk.Id, original[stateIndex].Id, StringComparison.Ordinal)) continue;
                if (match.HasValue) return null;
                match = original[stateIndex];
            }
            return match;
        }
        return null;
    }

    private static void RebuildPageBlocks(
        IReadOnlyList<OfficeDocumentPage>? pages,
        IReadOnlyList<OfficeDocumentBlock> blocks,
        CancellationToken cancellationToken) {
        foreach (OfficeDocumentPage page in pages ?? Array.Empty<OfficeDocumentPage>()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (page == null) continue;
            page.Blocks = blocks.Where(block => BelongsToPage(block.Location, page)).ToArray();
        }
    }

    private static bool BelongsToPage(ReaderLocation location, OfficeDocumentPage page) {
        string? containerKind = page.Location?.SourceBlockKind;
        if (string.Equals(containerKind, "slide", StringComparison.OrdinalIgnoreCase)) {
            return location.Slide == (page.Location?.Slide ?? page.Number);
        }
        if (string.Equals(containerKind, "sheet", StringComparison.OrdinalIgnoreCase)) {
            string? sheet = page.Location?.Sheet ?? page.Name;
            return string.Equals(location.Sheet, sheet, StringComparison.Ordinal);
        }
        return location.Page == (page.Location?.Page ?? page.Number);
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

internal sealed class ProcessedChunkAggregateSnapshot {
    internal ProcessedChunkAggregateSnapshot(
        OfficeDocumentReadResult document,
        bool chunkDerived,
        string? markdown,
        IReadOnlyList<ProcessedChunkState> chunks,
        IReadOnlyList<ProcessedBlockState> blocks) {
        Document = document;
        ChunkDerived = chunkDerived;
        Markdown = markdown;
        Chunks = chunks;
        Blocks = blocks;
    }

    internal OfficeDocumentReadResult Document { get; }
    internal bool ChunkDerived { get; }
    internal string? Markdown { get; }
    internal IReadOnlyList<ProcessedChunkState> Chunks { get; }
    internal IReadOnlyList<ProcessedBlockState> Blocks { get; }
}

internal readonly struct ProcessedChunkState {
    internal ProcessedChunkState(
        ReaderChunk? chunk,
        string? id,
        string? text,
        string? markdown,
        bool continuesPreviousChunk) {
        Chunk = chunk;
        Id = id;
        Text = text;
        Markdown = markdown;
        ContinuesPreviousChunk = continuesPreviousChunk;
    }

    internal ReaderChunk? Chunk { get; }
    internal string? Id { get; }
    internal string? Text { get; }
    internal string? Markdown { get; }
    internal bool ContinuesPreviousChunk { get; }
}

internal readonly struct ProcessedBlockState {
    internal ProcessedBlockState(
        OfficeDocumentBlock? block,
        string? id,
        string? kind,
        string? text,
        ReaderLocation? location) {
        Block = block;
        Id = id;
        Kind = kind;
        Text = text;
        Location = location;
    }

    internal OfficeDocumentBlock? Block { get; }
    internal string? Id { get; }
    internal string? Kind { get; }
    internal string? Text { get; }
    internal ReaderLocation? Location { get; }
}
