using System.Text;

namespace OfficeIMO.Reader;

/// <summary>Plain text and citation metadata for one page-like container.</summary>
public sealed class OfficeDocumentPageText {
    internal OfficeDocumentPageText(OfficeDocumentPage page, int totalPageCount, OfficeDocumentPageProvenance provenance, string text) {
        Page = page; TotalPageCount = totalPageCount; Provenance = provenance; Text = text;
    }
    /// <summary>Source page, slide or worksheet.</summary>
    public OfficeDocumentPage Page { get; }
    /// <summary>Total physical page count known to the read operation.</summary>
    public int TotalPageCount { get; }
    /// <summary>How the boundary was obtained.</summary>
    public OfficeDocumentPageProvenance Provenance { get; }
    /// <summary>Plain text in canonical source order, with tab-separated table cells.</summary>
    public string Text { get; }
}

public static partial class OfficeDocumentReadResultExtensions {
    /// <summary>Projects canonical blocks and tables into plain text without rereading the source.</summary>
    public static string GetText(this OfficeDocumentReadResult document, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        cancellationToken.ThrowIfCancellationRequested();
        var text = new StringBuilder();
        foreach (TextSegment segment in EnumerateTextSegments(document, cancellationToken)) {
            if (text.Length > 0 && !segment.Continues) text.AppendLine();
            text.Append(segment.Text);
        }
        return text.ToString().TrimEnd();
    }

    /// <summary>One nonempty piece of <see cref="GetText"/> output: a chunk, block or table line, in output order.</summary>
    internal readonly struct TextSegment {
        internal TextSegment(string text, bool continues, OfficeDocumentContentItem item) {
            Text = text; Continues = continues; Item = item;
        }
        internal string Text { get; }
        /// <summary>Appended without a line break, continuing the previous segment.</summary>
        internal bool Continues { get; }
        internal OfficeDocumentContentItem Item { get; }
    }

    /// <summary>The segments <see cref="GetText"/> joins, shared with the text projection so both produce the same text.</summary>
    internal static IEnumerable<TextSegment> EnumerateTextSegments(OfficeDocumentReadResult document, CancellationToken cancellationToken) {
        var content = document.EnumerateContent().ToArray();
        var tableLocations = content.Where(item => item.Table != null && !string.IsNullOrWhiteSpace(item.Location?.BlockAnchor))
            .GroupBy(item => item.Location!.BlockAnchor!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Select(item => item.Location!).ToArray(), StringComparer.Ordinal);
        bool EmitsBlock(OfficeDocumentContentItem item) => item.Block != null && !HasMatchingStructuredTable(item, tableLocations);
        var emittedChunks = content.Where(item => item.Chunk != null || EmitsBlock(item))
            .Select(item => item.Chunk ?? item.Block?.ChunkProjection).OfType<ReaderChunk>();
        var chunkTables = OfficeDocumentModelTraversal.TablesCoveredByTextChunks(emittedChunks, content);
        foreach (var item in content) {
            cancellationToken.ThrowIfCancellationRequested();
            if (item.Chunk != null) {
                if (!string.IsNullOrEmpty(item.Chunk.Text)) yield return new TextSegment(item.Chunk.Text, item.Chunk.ContinuesPreviousChunk, item);
            } else if (item.Table is ReaderTable table && !chunkTables.Contains(table)) {
                if (table.Columns.Count > 0) {
                    string columns = string.Join("\t", table.Columns);
                    if (columns.Length > 0) yield return new TextSegment(columns, false, item);
                }
                foreach (var row in table.Rows) {
                    cancellationToken.ThrowIfCancellationRequested();
                    string line = string.Join("\t", row);
                    if (line.Length > 0) yield return new TextSegment(line, false, item);
                }
            } else if (EmitsBlock(item) && !string.IsNullOrEmpty(item.Block!.Text)) {
                yield return new TextSegment(item.Block.Text, item.Block.ChunkProjection?.ContinuesPreviousChunk == true, item);
            }
        }
    }

    /// <summary>Projects each existing page-like container into plain text with its original citation metadata.</summary>
    public static IReadOnlyList<OfficeDocumentPageText> GetPageText(this OfficeDocumentReadResult document, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        cancellationToken.ThrowIfCancellationRequested();
        int total = document.GetTotalPageCount();
        var provenance = document.GetPageProvenance();
        var pages = new List<OfficeDocumentPageText>();
        foreach (var page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            cancellationToken.ThrowIfCancellationRequested();
            pages.Add(new OfficeDocumentPageText(page, total, provenance,
                new OfficeDocumentReadResult { Pages = new[] { page } }.GetText(cancellationToken)));
        }
        return pages.AsReadOnly();
    }
}
