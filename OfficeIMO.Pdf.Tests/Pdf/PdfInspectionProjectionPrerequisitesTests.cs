using System.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfInspectionProjectionPrerequisitesTests {
    [Fact]
    public void OptionalImagesAndWidgetsAreSkippedBeforeMaterializationWhileNativeTextRemains() {
        byte[] bytes = PdfInspectionProjectionFixture.Build(32);
        PdfDocument bounded = PdfDocument.Load(bytes, new PdfLoadOptions { Limits = new PdfReadLimits { MaxContentOperands = 20 } });
        Assert.Throws<PdfReadLimitException>(() => bounded.ReadForInspection());
        var noImages = new UnexpectedImageStage();
        PdfLoadOptions options = new PdfLoadOptions { IncludeImages = false, IncludeFormWidgets = false,
            Pipeline = new PdfUnderstandingPipelineOptions { ImageRegionDetection = noImages } };
        PdfInspectionReadResult result = bounded.ReadForInspection(options);
        Assert.True(result.CompletedAllSelectedPages); Assert.False(result.Document.IncludesImages); Assert.False(result.Document.IncludesFormWidgets);
        Assert.Contains("Native page text", result.Document.Text); Assert.Contains("Nested form text", result.Document.Text);
        Assert.Empty(result.Document.Images); Assert.Empty(result.Document.FormWidgets);
        Assert.False(Assert.Single(result.Document.Pages).HasVisualPrimitiveCounts);
        Assert.NotEmpty(Assert.Single(result.Document.Pages).Analysis.DecodedRuns);
        Assert.Single(result.SourceInfo.FormFields); Assert.Single(result.SourceInfo.Pages[0].FormWidgets);
        Assert.Equal("Source value", result.SourceInfo.FormFields[0].Value);
        Assert.False(options.Clone().IncludeImages); Assert.False(options.Clone().IncludeFormWidgets);
        PdfDocumentReadResult ordinary = PdfDocument.Load(bytes).Read();
        Assert.True(ordinary.IncludesImages); Assert.True(ordinary.IncludesFormWidgets);
        Assert.True(Assert.Single(ordinary.Pages).HasVisualPrimitiveCounts);
        Assert.NotEmpty(ordinary.Images); Assert.NotEmpty(ordinary.FormWidgets);
    }

    [Fact]
    public void TextProjectionDoesNotTurnTextLimitsOrCancellationIntoPartialSuccess() {
        byte[] bytes = PdfInspectionProjectionFixture.Build();
        var options = new PdfLoadOptions { IncludeImages = false, IncludeFormWidgets = false,
            Pipeline = new PdfUnderstandingPipelineOptions { MaxTextCharactersPerPage = 1 } };
        Assert.Throws<PdfReadLimitException>(() => PdfDocument.Load(bytes).ReadForInspection(options));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfDocument.Load(bytes).ReadForInspection(
            new PdfLoadOptions { IncludeImages = false }, cancellation.Token));
    }

    private sealed class UnexpectedImageStage : IPdfImageRegionDetectionStage {
        public IReadOnlyList<PdfUnderstandingImageRegion> Detect(PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingSemanticElement> semanticElements) => throw new InvalidOperationException("Unrequested image stage was invoked.");
    }
}
