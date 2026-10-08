namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Inline pictures stay upright while their advance and cross-axis size follow the cell's turned text axes.</summary>
    private static IReadOnlyList<PdfTextRun> ProjectTableCellInlineImages(IReadOnlyList<PdfTextRun> runs, int rotation) {
        if (rotation == 0) return runs;
        List<PdfTextRun>? projected = null;
        for (int i = 0; i < runs.Count; i++) {
            if (runs[i].InlineElement is not PdfInlineImage image) continue;
            projected ??= new List<PdfTextRun>(runs);
            projected[i] = runs[i].WithInlineElement(image.ForTableTextRotation(rotation));
        }
        return projected ?? runs;
    }

    private static IReadOnlyList<PdfTableCellParagraph> ProjectTableCellInlineImageParagraphs(IReadOnlyList<PdfTableCellParagraph> paragraphs, int rotation) {
        if (rotation == 0) return paragraphs;
        List<PdfTableCellParagraph>? projected = null;
        for (int i = 0; i < paragraphs.Count; i++) {
            PdfTableCellParagraph paragraph = paragraphs[i];
            IReadOnlyList<PdfTextRun> runs = ProjectTableCellInlineImages(paragraph.Runs, rotation);
            if (ReferenceEquals(runs, paragraph.Runs)) continue;
            projected ??= new List<PdfTableCellParagraph>(paragraphs);
            projected[i] = new PdfTableCellParagraph(runs, paragraph.SpacingAfter, paragraph.Align, paragraph.SpacingBefore,
                paragraph.LeftIndent, paragraph.RightIndent, paragraph.FirstLineIndent, paragraph.LineHeight,
                paragraph.DefaultTabStopWidth, paragraph.TabStops, paragraph.FontSize, paragraph.LineSpacing,
                paragraph.WidowControl, paragraph.KeepTogether, paragraph.KeepWithNext);
        }
        return projected ?? paragraphs;
    }
}
