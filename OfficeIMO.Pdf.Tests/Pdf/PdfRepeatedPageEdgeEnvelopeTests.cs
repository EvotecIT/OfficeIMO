using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRepeatedPageEdgeEnvelopeTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DriftingPageEdgesMeetTheDocumentQuorumWithinBoundedWork(bool footer) {
        const int pageCount = 100;
        PdfReadDocument document = CreateDocument(pageCount);
        int[] pageNumbers = Enumerable.Range(1, pageCount).ToArray();
        PdfUnderstandingPageResult[] pages = pageNumbers.Select(pageNumber => CreatePage(
            pageNumber,
            Element("Running operations", 40D + pageNumber / 10D, 200D + pageNumber / 10D,
                footer ? 50D : 950D),
            Element("Body " + pageNumber, 40D, 200D, 500D))).ToArray();

        // This is a caller-enforced work limit, not a timing or allocation budget.
        // Slight geometry drift must not require all-pairs comparisons across pages.
        IReadOnlyList<PdfUnderstandingPageResult> enriched = PdfDocumentSemanticEnricher.Enrich(
            document, pageNumbers, pages, 10, 4_000, CancellationToken.None);

        PdfUnderstandingSemanticKind expected = footer
            ? PdfUnderstandingSemanticKind.Footer : PdfUnderstandingSemanticKind.Header;
        Assert.Equal(pageNumbers, enriched.Select(page => page.PageNumber));
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++) {
            Assert.Equal(expected, enriched[pageIndex].Elements[0].Kind);
            Assert.Equal(PdfUnderstandingSemanticKind.Paragraph, enriched[pageIndex].Elements[1].Kind);
            Assert.Same(pages[pageIndex].Elements[0].Region, enriched[pageIndex].Elements[0].Region);
            Assert.Same(pages[pageIndex].Elements[1], enriched[pageIndex].Elements[1]);
        }
    }

    [Theory]
    [InlineData(80D, 240D, true)]
    [InlineData(80.001D, 240D, false)]
    [InlineData(80D, 240.001D, false)]
    public void PageEdgeGeometryThresholdsRemainInclusive(double left, double width, bool repeated) {
        PdfReadDocument document = CreateDocument(2);
        PdfUnderstandingPageResult[] pages = {
            CreatePage(1, Element("Running operations", 0D, 120D, 950D)),
            CreatePage(2, Element("Running operations", left, width, 950D))
        };

        IReadOnlyList<PdfUnderstandingPageResult> enriched = PdfDocumentSemanticEnricher.Enrich(
            document, new[] { 1, 2 }, pages, 10, 1_000, CancellationToken.None);

        Assert.All(enriched, page => Assert.Equal(
            repeated ? PdfUnderstandingSemanticKind.Header : PdfUnderstandingSemanticKind.Paragraph,
            Assert.Single(page.Elements).Kind));
    }

    [Fact]
    public void SeparatedGeometryClustersDoNotShareTheirPageQuorums() {
        PdfReadDocument document = CreateDocument(6);
        int[] pageNumbers = Enumerable.Range(1, 6).ToArray();
        double[] leftPositions = { 0D, 10D, 300D, 310D, 600D, 610D };
        PdfUnderstandingPageResult[] pages = pageNumbers.Select(pageNumber => CreatePage(
            pageNumber, Element("Running operations", leftPositions[pageNumber - 1], 120D, 950D))).ToArray();

        IReadOnlyList<PdfUnderstandingPageResult> enriched = PdfDocumentSemanticEnricher.Enrich(
            document, pageNumbers, pages, 10, 1_000, CancellationToken.None);

        // Six pages require three distinct pages in one compatible neighborhood.
        Assert.All(enriched, page => Assert.Equal(PdfUnderstandingSemanticKind.Paragraph,
            Assert.Single(page.Elements).Kind));
    }

    [Fact]
    public void RepeatedPageSelectionsDoNotEstablishAUniquePageQuorum() {
        PdfReadDocument document = CreateDocument(4);
        int[] pageNumbers = { 1, 1, 1, 2, 3, 4 };
        PdfUnderstandingPageResult[] pages = pageNumbers.Select(pageNumber => CreatePage(
            pageNumber, Element("Running operations", 40D, 200D,
                pageNumber == 1 ? 950D : 500D))).ToArray();

        IReadOnlyList<PdfUnderstandingPageResult> enriched = PdfDocumentSemanticEnricher.Enrich(
            document, pageNumbers, pages, 10, 1_000, CancellationToken.None);

        Assert.Equal(pageNumbers, enriched.Select(page => page.PageNumber));
        Assert.All(enriched, page => Assert.Equal(PdfUnderstandingSemanticKind.Paragraph,
            Assert.Single(page.Elements).Kind));
    }

    private static PdfReadDocument CreateDocument(int pageCount) {
        PdfDocument document = PdfDocument.Create(new PdfOptions { PageSize = new PageSize(1_000D, 1_000D) });
        for (int page = 0; page < pageCount; page++) {
            document.Paragraph(paragraph => paragraph.Text("Page placeholder"));
            if (page + 1 < pageCount) document.PageBreak();
        }
        return PdfReadDocument.Open(document.ToBytes());
    }

    private static PdfUnderstandingSemanticElement Element(string text, double left, double width, double baseline) {
        var word = new PdfUnderstandingWord(text, left, left + width, baseline, 10D, 0D,
            Array.Empty<PdfTextSpan>());
        var line = new PdfUnderstandingLine(new[] { word });
        return new PdfUnderstandingSemanticElement(new PdfUnderstandingRegion(new[] { line }),
            PdfUnderstandingSemanticKind.Paragraph);
    }

    private static PdfUnderstandingPageResult CreatePage(int pageNumber, params PdfUnderstandingSemanticElement[] elements) {
        PdfUnderstandingRegion[] regions = elements.Select(element => element.Region).ToArray();
        PdfUnderstandingLine[] lines = regions.SelectMany(region => region.Lines).ToArray();
        return new PdfUnderstandingPageResult(pageNumber, Array.Empty<PdfTextSpan>(),
            lines.SelectMany(line => line.Words).ToArray(), lines, regions, regions,
            Array.Empty<PdfReadingOrderEvidence>(), elements, Array.Empty<PdfUnderstandingStageTrace>());
    }
}
