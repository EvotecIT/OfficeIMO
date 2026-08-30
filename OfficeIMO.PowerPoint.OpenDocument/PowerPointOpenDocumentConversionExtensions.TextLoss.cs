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
        int count = RequiresDecorationApproximation(
            paragraph.UnderlineStyle,
            paragraph.UnderlineType,
            paragraph.LineThroughStyle) ? 1 : 0;
        foreach (OdpInlineNode node in paragraph.InlineNodes) {
            if (node.Run is OdpRun run &&
                RequiresDecorationApproximation(run.UnderlineStyle, run.UnderlineType, run.LineThroughStyle)) count++;
            if (node.Hyperlink is OdpHyperlink hyperlink &&
                RequiresDecorationApproximation(hyperlink.UnderlineStyle, hyperlink.UnderlineType, hyperlink.LineThroughStyle)) count++;
        }
        return count;
    }

    private static bool RequiresDecorationApproximation(
        OdfTextDecorationStyle? underlineStyle,
        OdfTextDecorationType? underlineType,
        OdfTextDecorationStyle? lineThroughStyle) =>
        lineThroughStyle is not (null or OdfTextDecorationStyle.None or OdfTextDecorationStyle.Solid) ||
        underlineType == OdfTextDecorationType.Double &&
        underlineStyle is not (null or OdfTextDecorationStyle.None or OdfTextDecorationStyle.Solid or OdfTextDecorationStyle.Wave);
}
