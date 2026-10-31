using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void RenderNotes(RtfDocument document, PdfCore.PdfDocument pdf, RtfToPdfOptions options, PdfRenderState state) {
        if (state.NoteReferences.Count == 0) {
            return;
        }

        if (!options.IncludeNotes) {
            AddConversionWarning(
                options,
                "NotesSkipped",
                "Notes",
                "RTF note bodies were skipped because IncludeNotes is false.",
                new Dictionary<string, string> {
                    ["Count"] = state.NoteReferences.Count.ToString(System.Globalization.CultureInfo.InvariantCulture)
                });
            return;
        }

        pdf.HR(spacingBefore: 8D, spacingAfter: 4D);
        foreach (PdfNoteReference noteReference in state.NoteReferences) {
            if (noteReference.Note.Paragraphs.Count == 0) {
                pdf.Paragraph(paragraph => paragraph.Runs(new[] {
                    state.CreateDefaultRun(GetNoteLabel(noteReference), bold: true, fontSize: global::OfficeIMO.OfficeLength.Points(9D)),
                    state.CreateDefaultRun(string.Empty, fontSize: global::OfficeIMO.OfficeLength.Points(9D))
                }));
            }

            for (int i = 0; i < noteReference.Note.Paragraphs.Count; i++) {
                RtfParagraph source = noteReference.Note.Paragraphs[i];
                RecordParagraphBarTabDiagnostic(document, source, options);
                if (RtfPdfMapping.HasPageBreakBefore(document, source)) pdf.PageBreak();
                var runs = new List<PdfCore.PdfTextRun>();
                if (i == 0) {
                    runs.Add(state.CreateDefaultRun(GetNoteLabel(noteReference), bold: true, fontSize: global::OfficeIMO.OfficeLength.Points(9D)));
                }

                int bodyStartIndex = runs.Count;
                AppendParagraphRuns(document, source, runs, options, state, collectNotes: false);
                EnsureParagraphLineBoxStrut(document, source, runs, state, bodyStartIndex);
                RtfParagraph formatting = document.GetParagraphFormatting(source);
                pdf.Paragraph(
                    paragraph => paragraph.Runs(runs),
                    RtfPdfMapping.ToPdfAlign(formatting.Alignment),
                    style: RtfPdfMapping.ToPdfParagraphStyle(document, source, runs, state.FontOptions));
            }
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
