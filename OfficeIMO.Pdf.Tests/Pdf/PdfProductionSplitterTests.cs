using System.Text;
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

    [Fact]
    public void Split_AcceptsTheExactCumulativeArtifactByteBudget() {
        byte[] source = PdfProductionWorkflowTestSupport.CreatePdf("one", "two");
        var baselineOptions = new PdfProductionSplitOptions { MaximumPagesPerPart = 1 };
        PdfProductionSplitResult baseline = PdfProductionSplitter.Split(source, baselineOptions);

        PdfProductionSplitResult exact = PdfProductionSplitter.Split(source, new PdfProductionSplitOptions {
            MaximumPagesPerPart = 1,
            MaximumCumulativeArtifactBytes = baseline.CumulativeArtifactBytes
        });

        Assert.Equal(baseline.CumulativeArtifactBytes, exact.CumulativeArtifactBytes);
        Assert.Equal(baseline.Parts.Select(static part => part.SizeBytes), exact.Parts.Select(static part => part.SizeBytes));
        Assert.Throws<InvalidOperationException>(() => PdfProductionSplitter.Split(source, new PdfProductionSplitOptions {
            MaximumPagesPerPart = 1,
            MaximumCumulativeArtifactBytes = baseline.CumulativeArtifactBytes - 1L
        }));
    }

    [Fact]
    public void Split_ReservesStructuralReadLimitsForGeneratedArtifacts() {
        byte[] source = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R >>\nendobj\n" +
            "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 5 >>\nstartxref\n0\n%%EOF\n");
        int sourceObjectCount = PdfReadDocument.Open(source).RawStructure().TotalObjectCount;
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxIndirectObjects = sourceObjectCount,
                MaxObjectCharacters = 100,
                MaxTokensPerObject = 100
            }
        };

        PdfProductionSplitResult result = PdfProductionSplitter.Split(
            source,
            new PdfProductionSplitOptions { MaximumPagesPerPart = 1 },
            readOptions);

        Assert.Single(result.Parts);
        Assert.All(result.Parts, static part => Assert.Single(part.ToDocument().Reader.Pages()));
    }
}
