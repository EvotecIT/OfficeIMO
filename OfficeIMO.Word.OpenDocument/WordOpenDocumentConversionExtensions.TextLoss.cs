using OfficeIMO.OpenDocument;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    private static int CountNonSolidTextDecorations(OdtDocument document) {
        int count = document.ContentBlocks.Sum(block => block.Paragraph != null
            ? CountNonSolidTextDecorations(block.Paragraph)
            : block.Table!.Rows.Sum(row => row.Cells.Sum(cell =>
                cell.Paragraphs.Sum(CountNonSolidTextDecorations))));
        count += document.PageLayout.Header.Paragraphs.Sum(CountNonSolidTextDecorations);
        count += document.PageLayout.Footer.Paragraphs.Sum(CountNonSolidTextDecorations);
        return count;
    }

    private static int CountNonSolidTextDecorations(OdtParagraph paragraph) {
        int count = paragraph.UsesNonSolidUnderlineStyle || paragraph.UsesNonSolidLineThroughStyle ? 1 : 0;
        foreach (OdtInlineNode node in paragraph.InlineNodes) {
            if (node.Span is OdtSpan span &&
                (span.UsesNonSolidUnderlineStyle || span.UsesNonSolidLineThroughStyle)) count++;
            if (node.Hyperlink is OdtHyperlink hyperlink &&
                (hyperlink.UsesNonSolidUnderlineStyle || hyperlink.UsesNonSolidLineThroughStyle)) count++;
        }
        return count;
    }
}
