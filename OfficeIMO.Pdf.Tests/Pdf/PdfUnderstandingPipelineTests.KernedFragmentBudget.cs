using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfUnderstandingPipelineTests {
    [Fact]
    public void KernedFragmentJoinsChargeRepeatedCopyWorkBeforeAllocation() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("placeholder")).ToBytes();
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(page, 1, new PdfTextLayoutOptions(),
            maxTextCharactersPerPage: 100_000, maxWordsPerPage: 100_000, maxWorkUnitsPerPage: 10_000);
        PdfTextSpan[] runs = Enumerable.Range(0, 300)
            .Select(index => new PdfTextSpan("aa", "F1", 12D, 40D + index * 2D, 700D, 2D,
                color: null, isVisible: true, rotationDegrees: 0D, baseFont: null, clipPath: null,
                glyphSequenceProgressesLeftToRight: true))
            .ToArray();

        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
            PdfAdvancedUnderstandingStages.WordGrouping.GroupWords(context, runs));
        Assert.Equal(PdfReadLimitKind.UnderstandingWork, error.Kind);
        Assert.Equal(10_000, error.Limit);

        var permissive = new PdfUnderstandingPageContext(page, 1, new PdfTextLayoutOptions(),
            maxTextCharactersPerPage: 100_000, maxWordsPerPage: 100_000, maxWorkUnitsPerPage: 1_000_000);
        PdfUnderstandingWord joined = Assert.Single(PdfAdvancedUnderstandingStages.WordGrouping.GroupWords(permissive, runs));
        Assert.Equal(300, joined.SourceRuns.Count);
        Assert.Equal(600, joined.Text.Length);
    }
}
