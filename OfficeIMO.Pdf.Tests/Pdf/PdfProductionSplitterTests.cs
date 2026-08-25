using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfProductionSplitterTests {
    [Fact]
    public void Split_CombinesPageCountAndContentBoundariesWithoutLosingPages() {
        byte[] source = PdfProductionWorkflowTestSupport.CreatePdf(
            "Batch one page one",
            "Batch one page two",
            "START RECORD two",
            "Record two page two",
            "Record two page three");

        PdfProductionSplitResult result = PdfProductionSplitter.Split(source, new PdfProductionSplitOptions {
            MaximumPagesPerPart = 2,
            BoundaryText = "start record"
        });

        Assert.Equal(5, result.SourcePageCount);
        Assert.Equal(new[] { new[] { 1, 2 }, new[] { 3, 4 }, new[] { 5 } }, result.Parts.Select(static part => part.SourcePages.ToArray()));
        Assert.Equal(new[] { PdfProductionSplitReason.PageCount, PdfProductionSplitReason.PageCount, PdfProductionSplitReason.EndOfDocument }, result.Parts.Select(static part => part.Reason));
        Assert.All(result.Parts, static part => Assert.Equal(part.SourcePages.Count, PdfInspector.Inspect(part.ToBytes()).PageCount));
    }

    [Fact]
    public void Split_UsesTargetSizeAndReportsAnOversizedSinglePage() {
        byte[] source = PdfProductionWorkflowTestSupport.CreatePdf("Size one", "Size two", "Size three");
        long onePageSize = PdfPageExtractor.ExtractPages(source, new[] { 1 }).LongLength;
        long twoPageSize = PdfPageExtractor.ExtractPages(source, new[] { 1, 2 }).LongLength;
        long target = Math.Max(onePageSize, twoPageSize - 1L);

        PdfProductionSplitResult result = PdfProductionSplitter.Split(source, new PdfProductionSplitOptions {
            TargetPartSizeBytes = target
        });

        Assert.True(result.Parts.Count >= 2);
        Assert.Equal(new[] { 1, 2, 3 }, result.Parts.SelectMany(static part => part.SourcePages));
        Assert.All(result.Parts.Where(static part => !part.ExceedsTargetSize), part => Assert.True(part.SizeBytes <= target));

        PdfProductionSplitResult tinyTarget = PdfProductionSplitter.Split(source, new PdfProductionSplitOptions {
            TargetPartSizeBytes = 1
        });
        Assert.Equal(3, tinyTarget.Parts.Count);
        Assert.All(tinyTarget.Parts, static part => Assert.True(part.ExceedsTargetSize));
    }

    [Fact]
    public void Split_FailsClosedWhenArtifactProbeBudgetIsExceeded() {
        byte[] source = PdfProductionWorkflowTestSupport.CreatePdf("one", "two", "three");

        Assert.Throws<InvalidOperationException>(() => PdfProductionSplitter.Split(source, new PdfProductionSplitOptions {
            TargetPartSizeBytes = 1,
            MaximumArtifactProbes = 1
        }));
    }

    [Fact]
    public void Split_FailsClosedWhenCumulativeArtifactByteBudgetIsExceeded() {
        byte[] source = PdfProductionWorkflowTestSupport.CreatePdf("one", "two");

        Assert.Throws<InvalidOperationException>(() => PdfProductionSplitter.Split(source, new PdfProductionSplitOptions {
            TargetPartSizeBytes = long.MaxValue,
            MaximumCumulativeArtifactBytes = 1
        }));
    }
}
