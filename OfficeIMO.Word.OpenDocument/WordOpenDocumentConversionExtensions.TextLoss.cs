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

    private static int CountUnsupportedWritingModes(OdtDocument document) {
        int count = document.ContentBlocks.Sum(block => block.Paragraph != null
            ? IsUnsupportedWritingMode(block.Paragraph.WritingMode) ? 1 : 0
            : block.Table!.Rows.Sum(row => row.Cells.Sum(cell =>
                cell.Paragraphs.Count(paragraph => IsUnsupportedWritingMode(paragraph.WritingMode)))));
        count += document.PageLayout.Header.Paragraphs.Count(paragraph =>
            IsUnsupportedWritingMode(paragraph.WritingMode));
        count += document.PageLayout.Footer.Paragraphs.Count(paragraph =>
            IsUnsupportedWritingMode(paragraph.WritingMode));
        return count;
    }

    private static bool IsUnsupportedWritingMode(string? writingMode) =>
        !string.IsNullOrWhiteSpace(writingMode)
        && !string.Equals(writingMode, "lr", StringComparison.OrdinalIgnoreCase)
        && !string.Equals(writingMode, "lr-tb", StringComparison.OrdinalIgnoreCase)
        && !string.Equals(writingMode, "rl", StringComparison.OrdinalIgnoreCase)
        && !string.Equals(writingMode, "rl-tb", StringComparison.OrdinalIgnoreCase);

    private static int CountNonSolidTextDecorations(OdtParagraph paragraph) {
        int count = RequiresDecorationApproximation(
            paragraph.UnderlineStyle,
            paragraph.UnderlineType,
            paragraph.LineThroughStyle) ? 1 : 0;
        foreach (OdtInlineNode node in paragraph.InlineNodes) {
            if (node.Span is OdtSpan span &&
                RequiresDecorationApproximation(span.UnderlineStyle, span.UnderlineType, span.LineThroughStyle)) count++;
            if (node.Hyperlink is OdtHyperlink hyperlink &&
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
