using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageCompositionTests {
    [Fact]
    public void ReversePagesUsesSharedReorderEngine() {
        byte[] source = BuildNumberedPages(6);

        byte[] reversed = PdfPageEditor.ReversePages(source);

        Assert.Equal(new[] { "Page 6", "Page 5", "Page 4", "Page 3", "Page 2", "Page 1" }, ReadPageText(reversed));
    }

    [Fact]
    public void RepeatPageRangesRepeatsTheComposedSelection() {
        byte[] source = BuildNumberedPages(5);

        byte[] repeated = PdfPageEditor.RepeatPageRanges(
            source,
            repetitions: 2,
            new PdfPageRange(1, 2),
            new PdfPageRange(5, 5));

        Assert.Equal(new[] { "Page 1", "Page 2", "Page 5", "Page 1", "Page 2", "Page 5" }, ReadPageText(repeated));
    }

    [Fact]
    public void InterleavePageRangesUsesRoundRobinOrder() {
        byte[] source = BuildNumberedPages(6);

        byte[] interleaved = PdfPageEditor.InterleavePageRanges(
            source,
            new PdfPageRange(1, 3),
            new PdfPageRange(4, 6));

        Assert.Equal(new[] { "Page 1", "Page 4", "Page 2", "Page 5", "Page 3", "Page 6" }, ReadPageText(interleaved));
    }

    [Fact]
    public void InterleavePageRangesContinuesUnevenSelections() {
        byte[] source = BuildNumberedPages(5);

        byte[] interleaved = PdfPageEditor.InterleavePageRanges(
            source,
            new PdfPageRange(1, 2),
            new PdfPageRange(3, 5));

        Assert.Equal(new[] { "Page 1", "Page 3", "Page 2", "Page 4", "Page 5" }, ReadPageText(interleaved));
    }

    [Fact]
    public void PageCompositionValidatesSelectionsAndCounts() {
        byte[] source = BuildNumberedPages(3);

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RepeatPages(source, 0, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.InterleavePageRanges(source, new PdfPageRange(1, 2)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.InterleavePageRanges(
            source,
            new PdfPageRange(1, 2),
            new PdfPageRange(3, 4)));
    }

    private static byte[] BuildNumberedPages(int pageCount) {
        PdfDocument document = PdfDocument.Create();
        for (int page = 1; page <= pageCount; page++) {
            document.Paragraph(paragraph => paragraph.Text("Page " + page));
            if (page < pageCount) {
                document.PageBreak();
            }
        }

        return document.ToBytes();
    }

    private static string[] ReadPageText(byte[] pdf) => PdfReadDocument.Open(pdf)
        .Pages
        .Select(static page => page.ExtractText().Trim())
        .ToArray();
}
