using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void RenderNotes(RtfDocument document, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        if (!options.IncludeNotes || state.NoteReferences.Count == 0) {
            return;
        }

        pdf.HR(spacingBefore: 8D, spacingAfter: 4D);
        foreach (PdfNoteReference noteReference in state.NoteReferences) {
            List<PdfCore.TextRun> runs = new List<PdfCore.TextRun> {
                PdfCore.TextRun.Bolded(GetNoteLabel(noteReference), fontSize: 9D)
            };

            if (noteReference.Note.Paragraphs.Count == 0) {
                runs.Add(PdfCore.TextRun.Normal(string.Empty, fontSize: 9D));
            }

            for (int i = 0; i < noteReference.Note.Paragraphs.Count; i++) {
                if (i > 0) {
                    runs.Add(PdfCore.TextRun.LineBreak());
                }

                AppendParagraphRuns(document, noteReference.Note.Paragraphs[i], runs, options, state, collectNotes: false);
            }

            pdf.Paragraph(paragraph => paragraph.Runs(runs));
        }
    }

    private static string GetNoteLabel(PdfNoteReference reference) {
        string marker = string.IsNullOrWhiteSpace(reference.Marker)
            ? reference.Ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : reference.Marker.Trim();

        switch (reference.Note.Kind) {
            case RtfNoteKind.Endnote:
                return "Endnote " + marker + ": ";
            case RtfNoteKind.Annotation:
                string author = string.IsNullOrWhiteSpace(reference.Note.Author)
                    ? string.Empty
                    : " (" + reference.Note.Author!.Trim() + ")";
                return "Annotation " + marker + author + ": ";
            default:
                return "Footnote " + marker + ": ";
        }
    }
}
