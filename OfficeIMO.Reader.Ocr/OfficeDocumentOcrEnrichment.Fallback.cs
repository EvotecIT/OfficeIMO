namespace OfficeIMO.Reader;

public static partial class OfficeDocumentOcrEnrichmentExtensions {
    // Canonical content uses chunks only until the first text block exists. Capture that
    // fallback before adding OCR blocks, without changing or discarding the original chunks.
    internal static IEnumerable<ReaderChunk> FallbackChunks(OfficeDocumentReadResult document) =>
        document.EnumerateBlocks().Any(block => !string.IsNullOrWhiteSpace(block.Text))
            ? Enumerable.Empty<ReaderChunk>()
            : (document.Chunks ?? Array.Empty<ReaderChunk>()).Where(chunk => chunk != null && !string.IsNullOrWhiteSpace(chunk.Text));

    internal static IReadOnlyList<OfficeDocumentBlock> PreserveFallbackBlocks(OfficeDocumentReadResult document) {
        var blocks = new List<OfficeDocumentBlock>(document.Blocks ?? Array.Empty<OfficeDocumentBlock>());
        foreach (ReaderChunk chunk in FallbackChunks(document)) {
            blocks.Add(new OfficeDocumentBlock {
                Id = chunk.Id, Kind = "chunk", Text = chunk.Text, Location = CloneLocation(chunk.Location)
            });
        }
        return blocks;
    }
}
