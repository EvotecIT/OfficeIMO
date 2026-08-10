using OfficeIMO.OpenDocument;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static int CountNonSolidTextDecorations(OdpPresentation presentation) {
        int count = 0;
        foreach (OdpSlide slide in presentation.Slides) {
            foreach (OdpShape shape in slide.Shapes) {
                if (shape is OdpTextBox textBox) {
                    count += textBox.Paragraphs.Sum(CountNonSolidTextDecorations);
                } else if (shape is OdpTable table) {
                    count += table.Rows.Sum(row => row.Cells.Sum(cell =>
                        cell.Paragraphs.Sum(CountNonSolidTextDecorations)));
                }
            }
            if (slide.SpeakerNotes != null) {
                count += slide.SpeakerNotes.Paragraphs.Sum(CountNonSolidTextDecorations);
            }
        }
        return count;
    }

    private static int CountNonSolidTextDecorations(OdpParagraph paragraph) {
        int count = paragraph.UsesNonSolidUnderlineStyle || paragraph.UsesNonSolidLineThroughStyle ? 1 : 0;
        foreach (OdpInlineNode node in paragraph.InlineNodes) {
            if (node.Run is OdpRun run &&
                (run.UsesNonSolidUnderlineStyle || run.UsesNonSolidLineThroughStyle)) count++;
            if (node.Hyperlink is OdpHyperlink hyperlink &&
                (hyperlink.UsesNonSolidUnderlineStyle || hyperlink.UsesNonSolidLineThroughStyle)) count++;
        }
        return count;
    }
}
