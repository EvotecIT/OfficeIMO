namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrEnrichmentExtensions {
    // Canonical content uses chunks only until the first text block exists. Capture that
    // fallback before adding OCR blocks, without changing or discarding the original chunks.
    internal static IEnumerable<(ReaderChunk Chunk, int Index)> FallbackChunks(OfficeDocumentReadResult document) {
        if (document.EnumerateBlocks().Any(block => !string.IsNullOrWhiteSpace(block.Text))) yield break;
        IReadOnlyList<ReaderChunk> chunks = document.Chunks ?? Array.Empty<ReaderChunk>();
        for (int index = 0; index < chunks.Count; index++) {
            ReaderChunk chunk = chunks[index];
            if (chunk != null && !string.IsNullOrWhiteSpace(chunk.Text)) yield return (chunk, index);
        }
    }

    internal static IReadOnlyList<OfficeDocumentBlock> PreserveFallbackBlocks(OfficeDocumentReadResult document) {
        var blocks = new List<OfficeDocumentBlock>(document.Blocks ?? Array.Empty<OfficeDocumentBlock>());
        foreach (var (chunk, index) in FallbackChunks(document)) {
            blocks.Add(new OfficeDocumentBlock {
                Id = OfficeDocumentModelTraversal.BuildFallbackChunkBlockId(chunk, index),
                Kind = "chunk", Text = chunk.Text, Location = CloneLocation(chunk.Location)
            });
        }
        return blocks;
    }
}
