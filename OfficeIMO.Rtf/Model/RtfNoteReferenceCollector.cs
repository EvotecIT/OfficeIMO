namespace OfficeIMO.Rtf;

/// <summary>Collects note references across the complete semantic block and inline tree.</summary>
internal static class RtfNoteReferenceCollector {
    public static HashSet<RtfNote> Collect(RtfDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var notes = new HashSet<RtfNote>();
        CollectBlocks(document.Blocks, notes);
        foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) {
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) CollectParagraph(paragraph, notes);
        }

        return notes;
    }

    private static void CollectBlocks(IEnumerable<IRtfBlock> blocks, ISet<RtfNote> notes) {
        foreach (IRtfBlock block in blocks) {
            switch (block) {
                case RtfParagraph paragraph:
                    CollectParagraph(paragraph, notes);
                    break;
                case RtfTable table:
                    foreach (RtfTableRow row in table.Rows) {
                        foreach (RtfTableCell cell in row.Cells) CollectBlocks(cell.Blocks, notes);
                    }
                    break;
                case RtfObject rtfObject:
                    CollectParagraph(rtfObject.Result, notes);
                    break;
                case RtfShape shape:
                    foreach (RtfParagraph paragraph in shape.TextBoxParagraphs) CollectParagraph(paragraph, notes);
                    break;
            }
        }
    }

    private static void CollectParagraph(RtfParagraph paragraph, ISet<RtfNote> notes) {
        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run when run.Note != null:
                    notes.Add(run.Note);
                    break;
                case RtfGeneratedText generatedText when generatedText.Note != null:
                    notes.Add(generatedText.Note);
                    break;
                case RtfField field:
                    CollectParagraph(field.Result, notes);
                    break;
                case RtfObject rtfObject:
                    CollectParagraph(rtfObject.Result, notes);
                    break;
                case RtfShape shape:
                    foreach (RtfParagraph shapeParagraph in shape.TextBoxParagraphs) CollectParagraph(shapeParagraph, notes);
                    break;
            }
        }
    }
}
