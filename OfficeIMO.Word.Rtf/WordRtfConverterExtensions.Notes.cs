using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void AppendNote(WordParagraph wordRun, RtfNote note, RtfDocument? rtfDocument, RtfWordNumberingPlan numbering) {
        if (note.Kind == RtfNoteKind.Annotation) {
            AppendAnnotationComment(wordRun, note, rtfDocument, numbering);
            return;
        }

        if (note.Kind != RtfNoteKind.Footnote && note.Kind != RtfNoteKind.Endnote) {
            return;
        }

        RtfNoteKind projectedKind = ResolveNoteProjection(rtfDocument, wordRun)?.ProjectKind(note.Kind) ?? note.Kind;
        RtfParagraph? firstParagraph = note.Paragraphs.FirstOrDefault();
        WordParagraph noteParagraph = firstParagraph != null
            ? CreateDetachedWordParagraph(wordRun._document, firstParagraph, rtfDocument, numbering)
            : new WordParagraph(wordRun._document, newParagraph: true, newRun: true);
        WordParagraph referenceRun = projectedKind == RtfNoteKind.Endnote
            ? WordEndNote.AddEndNote(wordRun._document, wordRun, noteParagraph).Paragraph
            : WordFootNote.AddFootNote(wordRun._document, wordRun, noteParagraph).Paragraph;
        AppendAdditionalNoteParagraphs(referenceRun, note, projectedKind, rtfDocument, numbering);
    }

    private static WordParagraph CreateDetachedWordParagraph(WordDocument document, RtfParagraph paragraph, RtfDocument? rtfDocument, RtfWordNumberingPlan numbering) {
        var wordParagraph = new WordParagraph(document, newParagraph: true, newRun: false);
        if (rtfDocument != null) {
            ApplyParagraphFormatting(wordParagraph, paragraph, rtfDocument, numbering);
        }

        AppendRuns(wordParagraph, paragraph, rtfDocument, numbering);
        return wordParagraph;
    }

    private static void AppendAdditionalNoteParagraphs(WordParagraph referenceRun, RtfNote note, RtfNoteKind projectedKind, RtfDocument? rtfDocument, RtfWordNumberingPlan numbering) {
        if (note.Paragraphs.Count <= 1) {
            return;
        }

        long? referenceId = projectedKind == RtfNoteKind.Endnote
            ? referenceRun.EndNote?.ReferenceId
            : referenceRun.FootNote?.ReferenceId;
        if (!referenceId.HasValue) {
            return;
        }

        OpenXmlCompositeElement? noteElement = projectedKind == RtfNoteKind.Endnote
            ? referenceRun._document._wordprocessingDocument.MainDocumentPart?.EndnotesPart?.Endnotes?
                .ChildElements.OfType<Endnote>()
                .FirstOrDefault(item => item.Id?.Value == referenceId.Value)
            : referenceRun._document._wordprocessingDocument.MainDocumentPart?.FootnotesPart?.Footnotes?
                .ChildElements.OfType<Footnote>()
                .FirstOrDefault(item => item.Id?.Value == referenceId.Value);
        if (noteElement == null) return;

        foreach (RtfParagraph paragraph in note.Paragraphs.Skip(1)) {
            WordParagraph wordParagraph = CreateDetachedWordParagraph(referenceRun._document, paragraph, rtfDocument, numbering);
            noteElement.Append(wordParagraph._paragraph);
        }
    }

}
