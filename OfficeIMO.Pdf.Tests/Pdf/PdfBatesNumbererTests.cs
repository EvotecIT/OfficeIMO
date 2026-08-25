using OfficeIMO.Pdf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfBatesNumbererTests {
    [Fact]
    public void Apply_NumbersSelectedPagesContinuouslyAcrossDocuments() {
        byte[] first = PdfProductionWorkflowTestSupport.CreatePdf("First one", "First two");
        byte[] second = PdfProductionWorkflowTestSupport.CreatePdf("Second one", "Second two");
        var secondInput = new PdfBatesDocument(second, "second.pdf") {
            TargetPages = PdfPageSelector.Parse("2")
        };

        PdfBatesBatchResult result = PdfBatesNumberer.Apply(
            new[] { new PdfBatesDocument(first, "first.pdf"), secondInput },
            new PdfBatesNumberingOptions {
                StartNumber = 42,
                Prefix = "CASE-",
                Suffix = "-EV",
                MinimumDigits = 4,
                Position = PdfBatesPosition.TopCenter
            });

        Assert.Equal(new long[] { 42, 43, 44 }, result.Assignments.Select(static assignment => assignment.Number));
        Assert.Equal(new[] { "CASE-0042-EV", "CASE-0043-EV", "CASE-0044-EV" }, result.Assignments.Select(static assignment => assignment.Text));
        Assert.Equal(45, result.NextNumber);
        Assert.All(result.Documents, static document => Assert.True(document.Preservation.IsPreserved));

        string[] firstPages = PdfProductionWorkflowTestSupport.ReadPageTexts(result.Documents[0].ToBytes());
        string[] secondPages = PdfProductionWorkflowTestSupport.ReadPageTexts(result.Documents[1].ToBytes());
        Assert.Contains("CASE-0042-EV", firstPages[0], StringComparison.Ordinal);
        Assert.Contains("CASE-0043-EV", firstPages[1], StringComparison.Ordinal);
        Assert.DoesNotContain("CASE-", secondPages[0], StringComparison.Ordinal);
        Assert.Contains("CASE-0044-EV", secondPages[1], StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_RejectsDuplicatePageSelectionsBeforeAssigningNumbers() {
        var input = new PdfBatesDocument(PdfProductionWorkflowTestSupport.CreatePdf("Duplicate selection")) {
            TargetPages = PdfPageSelector.Parse("1,1")
        };

        Assert.Throws<ArgumentException>(() => PdfBatesNumberer.Apply(new[] { input }));
    }

    [Fact]
    public void Apply_RejectsUndefinedPositions() {
        byte[] source = PdfProductionWorkflowTestSupport.CreatePdf("Undefined position");

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfBatesNumberer.Apply(
            new[] { new PdfBatesDocument(source) },
            new PdfBatesNumberingOptions { Position = (PdfBatesPosition)99 }));
    }

    [Fact]
    public void Apply_MapsTopAndBottomPositionsToTopLeftCanvasCoordinates() {
        byte[] source = PdfProductionWorkflowTestSupport.CreatePdf("Body marker");
        byte[] top = PdfBatesNumberer.Apply(
            new[] { new PdfBatesDocument(source) },
            new PdfBatesNumberingOptions { Prefix = "TOP-", MinimumDigits = 2, Position = PdfBatesPosition.TopLeft }).Documents[0].ToBytes();
        byte[] bottom = PdfBatesNumberer.Apply(
            new[] { new PdfBatesDocument(source) },
            new PdfBatesNumberingOptions { Prefix = "BOTTOM-", MinimumDigits = 2, Position = PdfBatesPosition.BottomLeft }).Documents[0].ToBytes();

        PdfLogicalTextBlock topBlock = Assert.Single(PdfLogicalDocument.Load(top).TextBlocks, static block => block.Text.Contains("TOP-01", StringComparison.Ordinal));
        PdfLogicalTextBlock bottomBlock = Assert.Single(PdfLogicalDocument.Load(bottom).TextBlocks, static block => block.Text.Contains("BOTTOM-01", StringComparison.Ordinal));
        Assert.True(topBlock.BaselineY > bottomBlock.BaselineY);
    }

    [Fact]
    public void Apply_RejectsLabelsThatCannotFitTheConfiguredRectangle() {
        byte[] source = PdfProductionWorkflowTestSupport.CreatePdf("Body marker");

        InvalidOperationException widthError = Assert.Throws<InvalidOperationException>(() => PdfBatesNumberer.Apply(
            new[] { new PdfBatesDocument(source) },
            new PdfBatesNumberingOptions { Prefix = new string('W', 200) }));
        Assert.Contains("does not fit", widthError.Message, StringComparison.Ordinal);

        InvalidOperationException heightError = Assert.Throws<InvalidOperationException>(() => PdfBatesNumberer.Apply(
            new[] { new PdfBatesDocument(source) },
            new PdfBatesNumberingOptions { Height = 5D, FontSize = 10D }));
        Assert.Contains("does not fit", heightError.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ReservesStructuralReadLimitsForGeneratedBatesObjects() {
        byte[] source = PdfProductionWorkflowTestSupport.CreatePdf("Tight Bates budget");
        int sourceObjectCount = PdfReadDocument.Open(source).RawStructure().TotalObjectCount;
        var input = new PdfBatesDocument(source) {
            ReadOptions = new PdfReadOptions {
                Limits = new PdfReadLimits { MaxIndirectObjects = sourceObjectCount }
            }
        };

        PdfBatesDocumentResult result = Assert.Single(PdfBatesNumberer.Apply(
            new[] { input },
            new PdfBatesNumberingOptions { Prefix = "7 0 obj-" }).Documents);

        Assert.Contains("7 0 obj-000001", result.ToDocument().Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ReservesContentReadLimitsForGeneratedBatesStreams() {
        byte[] source = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R >>\nendobj\n" +
            "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 5 >>\nstartxref\n0\n%%EOF\n");
        var input = new PdfBatesDocument(source) {
            ReadOptions = new PdfReadOptions {
                Limits = new PdfReadLimits {
                    MaxRawStreamBytes = 1,
                    MaxDecodedStreamBytes = 1,
                    MaxTotalDecodedStreamBytes = 1,
                    MaxPageContentBytes = 1,
                    MaxRetainedContentBytes = 1,
                    MaxDecodedTextCharacters = 1,
                    MaxObjectCharacters = 100,
                    MaxTokensPerObject = 50,
                    MaxObjectNestingDepth = 4,
                    MaxRevisions = 1,
                    MaxContentOperations = 1,
                    MaxContentOperands = 1,
                    MaxContentNestingDepth = 1
                }
            }
        };

        PdfBatesDocumentResult result = Assert.Single(PdfBatesNumberer.Apply(
            new[] { input },
            new PdfBatesNumberingOptions { Prefix = "startxref 123-" }).Documents);

        Assert.Contains("startxref 123-000001", result.ToDocument().Read.Text(), StringComparison.Ordinal);
    }
}
