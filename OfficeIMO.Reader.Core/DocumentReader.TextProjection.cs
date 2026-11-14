using System.Globalization;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>Builds the text projection's envelope from chunks a handler read directly as text.</summary>
    internal static OfficeDocumentReadResult CreateTextResult(IReadOnlyList<ReaderChunk> chunks, ReaderInputKind kind, IEnumerable<string> capabilities) => new() {
        Kind = kind,
        Chunks = chunks,
        CapabilitiesUsed = BuildChunkDocumentCapabilities(kind).Concat(capabilities)
            .Where(static value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal).ToArray()
    };

    /// <summary>
    /// Reduces a text-and-metadata result to the text projection: one chunk per <c>GetText()</c> segment, in output
    /// order, so <c>GetText()</c> over the reduced result returns the same text. Everything else except the source,
    /// capabilities and diagnostics is cleared. A result that already holds only plain chunks keeps them.
    /// </summary>
    internal static void ReduceToText(OfficeDocumentReadResult result) {
        IReadOnlyList<ReaderChunk> source = result.Chunks ?? Array.Empty<ReaderChunk>();
        bool chunksOnly = (result.Blocks?.Count ?? 0) == 0 && (result.Tables?.Count ?? 0) == 0 && (result.Pages?.Count ?? 0) == 0
            && source.All(static chunk => chunk == null || chunk.Tables == null || chunk.Tables.Count == 0);
        result.Chunks = chunksOnly ? StripChunks(source) : BuildTextChunks(result);
        result.Markdown = null;
        result.Html = null;
        result.Json = null;
        result.Metadata = Array.Empty<OfficeDocumentMetadataEntry>();
        result.Pages = Array.Empty<OfficeDocumentPage>();
        result.Blocks = Array.Empty<OfficeDocumentBlock>();
        result.Tables = Array.Empty<ReaderTable>();
        result.Assets = Array.Empty<OfficeDocumentAsset>();
        result.Links = Array.Empty<OfficeDocumentLink>();
        result.Forms = Array.Empty<OfficeDocumentFormField>();
        result.OcrCandidates = Array.Empty<OfficeDocumentOcrCandidate>();
        result.Visuals = Array.Empty<ReaderVisual>();
        result.NestedDocuments = Array.Empty<OfficeDocumentNestedResult>();
    }

    private static IReadOnlyList<ReaderChunk> StripChunks(IReadOnlyList<ReaderChunk> chunks) {
        var stripped = new ReaderChunk[chunks.Count];
        for (int index = 0; index < chunks.Count; index++) {
            ReaderChunk chunk = chunks[index];
            stripped[index] = chunk == null || IsPlainText(chunk) ? chunk! : StripChunk(chunk, chunk.Location);
        }
        return stripped;
    }

    private static IReadOnlyList<ReaderChunk> BuildTextChunks(OfficeDocumentReadResult result) {
        var chunks = new List<ReaderChunk>();
        foreach (OfficeDocumentReadResultExtensions.TextSegment segment in
                 OfficeDocumentReadResultExtensions.EnumerateTextSegments(result, CancellationToken.None)) {
            // GetText orders chunks by location; the output position becomes the logical order so the reduced
            // chunks keep exactly this order whatever containers the original items came from.
            ReaderLocation location = ReaderHierarchicalChunker.CloneLocation(segment.Item.Location);
            location.LogicalOrder = chunks.Count;
            ReaderChunk? origin = segment.Item.Chunk ?? segment.Item.Block?.ChunkProjection;
            if (origin != null && string.Equals(origin.Text, segment.Text, StringComparison.Ordinal)) {
                chunks.Add(StripChunk(origin, location));
                continue;
            }
            chunks.Add(new ReaderChunk {
                Id = !string.IsNullOrWhiteSpace(segment.Item.Block?.Id) && segment.Item.Table == null
                    ? segment.Item.Block!.Id
                    : "text:" + chunks.Count.ToString("D4", CultureInfo.InvariantCulture),
                Kind = result.Kind,
                Location = location,
                Text = segment.Text,
                ContinuesPreviousChunk = segment.Continues
            });
        }
        return chunks;
    }

    private static bool IsPlainText(ReaderChunk chunk) =>
        chunk.Markdown == null && chunk.Tables == null && chunk.Visuals == null && chunk.FormFields == null && chunk.Actions == null && chunk.Diagnostics == null;

    private static ReaderChunk StripChunk(ReaderChunk chunk, ReaderLocation location) {
        ReaderChunk copy = chunk.CopyForContainer();
        copy.Location = location;
        copy.Markdown = null;
        copy.Tables = null;
        copy.Visuals = null;
        copy.FormFields = null;
        copy.Actions = null;
        copy.Diagnostics = null;
        return copy;
    }
}
