using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

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
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        var text = new StringBuilder();
        void Append(string? value, bool continues = false) {
            if (string.IsNullOrEmpty(value)) return;
            if (text.Length > 0 && !continues) text.AppendLine();
            text.Append(value);
        }
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
                Append(item.Chunk.Text, item.Chunk.ContinuesPreviousChunk);
            } else if (item.Table is ReaderTable table && !chunkTables.Contains(table)) {
                if (table.Columns.Count > 0) Append(string.Join("\t", table.Columns));
                foreach (var row in table.Rows) {
                    cancellationToken.ThrowIfCancellationRequested();
                    Append(string.Join("\t", row));
                }
            } else if (EmitsBlock(item)) {
                Append(item.Block!.Text, item.Block.ChunkProjection?.ContinuesPreviousChunk == true);
            }
        }
        return text.ToString().TrimEnd();
    }

    /// <summary>Projects each existing page-like container into plain text with its original citation metadata.</summary>
    public static IReadOnlyList<OfficeDocumentPageText> GetPageText(this OfficeDocumentReadResult document, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
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
