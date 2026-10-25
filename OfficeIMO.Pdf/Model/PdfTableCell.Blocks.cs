namespace OfficeIMO.Pdf;

public sealed partial class PdfTableCell {
    /// <summary>Ordered captured cell content. Flat text remains the fast path when this list is empty.</summary>
    internal IReadOnlyList<IPdfBlock> Blocks { get; } = Array.Empty<IPdfBlock>();

    internal int BlockDepth { get; }

    private static IReadOnlyList<IPdfBlock> SnapshotBlocks(IEnumerable<IPdfBlock>? blocks) {
        if (blocks == null) return Array.Empty<IPdfBlock>();
        var snapshot = new List<IPdfBlock>();
        foreach (IPdfBlock block in blocks) {
            if (snapshot.Count >= 100_000) throw new ArgumentException("Cell content exceeded the 100,000-block limit.", nameof(blocks));
            switch (block) {
                case RichParagraphBlock paragraph:
                    snapshot.Add(new RichParagraphBlock(paragraph.Runs, paragraph.Align, paragraph.DefaultColor, paragraph.Style));
                    break;
                case TableBlock table:
                    if (table.CellBlockDepth >= 128) throw new ArgumentException("Nested table content exceeded the 128-level limit.", nameof(blocks));
                    snapshot.Add(table.Clone());
                    break;
                default:
                    throw new ArgumentException("Ordered table cell content supports captured paragraphs and inline tables.", nameof(blocks));
            }
        }
        return snapshot.AsReadOnly();
    }
}
