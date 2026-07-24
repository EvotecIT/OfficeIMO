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

    public static int CountHeaderFooterReferences(RtfDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        int count = 0;
        foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) {
            foreach (RtfParagraph paragraph in headerFooter.Paragraphs) {
                count += CountParagraphReferences(paragraph);
            }
        }

        return count;
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
                    CollectNote(run.Note, notes);
                    break;
                case RtfGeneratedText generatedText when generatedText.Note != null:
                    CollectNote(generatedText.Note, notes);
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

    private static void CollectNote(RtfNote note, ISet<RtfNote> notes) {
        if (!notes.Add(note)) return;
        foreach (RtfParagraph paragraph in note.Paragraphs) {
            CollectParagraph(paragraph, notes);
        }
    }

    private static int CountParagraphReferences(RtfParagraph paragraph) {
        int count = 0;
        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run when run.Note != null:
                    count += CountSerializedNoteSequence(run.Note);
                    break;
                case RtfGeneratedText generatedText when generatedText.Note != null:
                    count += CountSerializedNoteSequence(generatedText.Note);
                    break;
                case RtfField field:
                    count += CountParagraphReferences(field.Result);
                    break;
                case RtfObject rtfObject:
                    count += CountParagraphReferences(rtfObject.Result);
                    break;
                case RtfShape shape:
                    foreach (RtfParagraph shapeParagraph in shape.TextBoxParagraphs) {
                        count += CountParagraphReferences(shapeParagraph);
                    }
                    break;
            }
        }

        return count;
    }

    private static int CountSerializedNoteSequence(RtfNote note) =>
        CountSerializedNoteSequence(note, new HashSet<RtfNote>());

    private static int CountSerializedNoteSequence(RtfNote note, ISet<RtfNote> activeNotes) {
        if (!activeNotes.Add(note)) return 0;
        int count = 1;
        foreach (RtfParagraph paragraph in note.Paragraphs) {
            count += CountParagraphReferences(paragraph, activeNotes);
        }
        activeNotes.Remove(note);
        return count;
    }

    private static int CountParagraphReferences(RtfParagraph paragraph, ISet<RtfNote> activeNotes) {
        int count = 0;
        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run when run.Note != null:
                    count += CountSerializedNoteSequence(run.Note, activeNotes);
                    break;
                case RtfGeneratedText generatedText when generatedText.Note != null:
                    count += CountSerializedNoteSequence(generatedText.Note, activeNotes);
                    break;
                case RtfField field:
                    count += CountParagraphReferences(field.Result, activeNotes);
                    break;
                case RtfObject rtfObject:
                    count += CountParagraphReferences(rtfObject.Result, activeNotes);
                    break;
                case RtfShape shape:
                    foreach (RtfParagraph shapeParagraph in shape.TextBoxParagraphs) {
                        count += CountParagraphReferences(shapeParagraph, activeNotes);
                    }
                    break;
            }
        }
        return count;
    }
}
