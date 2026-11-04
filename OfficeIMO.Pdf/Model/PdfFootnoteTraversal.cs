namespace OfficeIMO.Pdf;

/// <summary>Visits authored references without materializing dynamic content or following definitions recursively.</summary>
internal static class PdfFootnoteTraversal {
    internal static IEnumerable<PdfFootnoteDefinition> References(IPdfBlock block) {
        IEnumerable<PdfTextRun> runs = block switch {
            RichParagraphBlock paragraph => paragraph.Runs,
            HeadingBlock heading => heading.Runs ?? Array.Empty<PdfTextRun>(),
            PdfListBlock list => list.RichItems.SelectMany(item => item.Runs),
            TableBlock table => table.Cells.SelectMany(row => row).SelectMany(cell => cell.Runs.Concat(cell.Paragraphs.SelectMany(paragraph => paragraph.Runs))),
            _ => Array.Empty<PdfTextRun>()
        };
        foreach (PdfTextRun run in runs)
            if (run.FootnoteReference is { } note) yield return note;
        foreach (IPdfBlock child in Children(block))
            foreach (PdfFootnoteDefinition note in References(child)) yield return note;
    }

    /// <summary>Authoring paragraphs inside a block, so flow that cannot place footnotes can reject later footnote runs.</summary>
    internal static IEnumerable<PdfParagraph> LiveParagraphs(IPdfBlock block) {
        if (block is RichParagraphBlock { Source: { } paragraph }) yield return paragraph;
        foreach (IPdfBlock child in Children(block))
            foreach (PdfParagraph nested in LiveParagraphs(child)) yield return nested;
    }

    private static IEnumerable<IPdfBlock> Children(IPdfBlock block) => block switch {
        FlowBlock { StaticBlocks: { } blocks } => blocks,
        PageBlock page => page.Blocks,
        ContainerBlock container => container.Blocks,
        MultiColumnBlock columns => columns.Blocks,
        RowBlock row => row.Columns.SelectMany(column => column.Blocks),
        SemanticBlock semantic => semantic.Blocks,
        LayerBlock layer => layer.Blocks,
        SectionBlock section => section.Blocks,
        TableBlock table => table.Cells.SelectMany(row => row).SelectMany(cell => cell.Blocks),
        _ => Array.Empty<IPdfBlock>()
    };
}
