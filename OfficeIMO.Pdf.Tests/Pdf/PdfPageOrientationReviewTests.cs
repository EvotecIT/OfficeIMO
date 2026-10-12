using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPageOrientationReviewTests {
    [Fact]
    public async Task ReviewIsExplicitSnapshotBoundAndAppliesOneRelativeBatch() {
        var document = Source().Pages.Rotate(90, 1);
        byte[] original = document.ToBytes();
        var options = new PdfOcrMergeOptions { Dpi = 72 };
        var review = await document.PrepareOrientationReviewAsync(Engine(90, .9), options);
        options.MinimumOrientationConfidence = 1;
        Assert.Equal(original, document.ToBytes());
        Assert.All(review.Pages, proposal => Assert.True(proposal.IsRecommended));
        Assert.Equal(90, review.Pages[0].OriginalRotationDegrees);
        var preview = await review.PreviewAsync(review.Pages[0]);
        Assert.NotEqual(preview.GetSourcePng(), preview.GetPreparedPng());
        var result = review.ApplyResult(document, new[] { review.Pages[0] });
        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));
        var updated = result.RequireValue();
        Assert.Equal(180, PdfReadDocument.Open(updated.ToBytes()).Pages[0].GetRotationDegrees());
        Assert.Equal(0, PdfReadDocument.Open(updated.ToBytes()).Pages[1].GetRotationDegrees());
        Assert.Throws<InvalidOperationException>(() => review.ApplyResult(updated, new[] { review.Pages[1] }));
        Assert.Equal(original, document.ToBytes());
    }

    [Theory]
    [InlineData(90, .1, true, false)]
    [InlineData(0, .9, false, false)]
    [InlineData(45, .9, false, false)]
    [InlineData(270, 2, false, false)]
    [InlineData(180, .9, true, true)]
    public async Task InvalidAndLowConfidenceEvidenceRequiresAnExplicitDecision(int correction, double confidence, bool applicable, bool recommended) {
        var document = Source(); byte[] before = document.ToBytes();
        var review = await document.PrepareOrientationReviewAsync(Engine(correction, confidence));
        var proposal = review.Pages[0];
        Assert.Equal(applicable, proposal.CanApply); Assert.Equal(recommended, proposal.IsRecommended);
        if (applicable) document = review.ApplyResult(document, new[] { proposal }).RequireValue();
        else Assert.Throws<ArgumentException>(() => review.ApplyResult(document, new[] { proposal }));
        Assert.Equal(applicable ? correction : 0, PdfReadDocument.Open(document.ToBytes()).Pages[0].GetRotationDegrees());
        if (!applicable) Assert.Equal(before, document.ToBytes());
    }

    [Fact]
    public async Task ForeignDuplicateAndCanceledSelectionsLeaveSourceUnchanged() {
        var document = Source(); byte[] before = document.ToBytes();
        var review = await document.PrepareOrientationReviewAsync(Engine(90, 1));
        var foreign = await document.PrepareOrientationReviewAsync(Engine(90, 1));
        Assert.Throws<ArgumentException>(() => review.ApplyResult(document, foreign.Pages));
        Assert.Throws<ArgumentException>(() => review.ApplyResult(document, new[] { review.Pages[0], review.Pages[0] }));
        Assert.ThrowsAny<OperationCanceledException>(() => review.ApplyResult(document, review.Pages, new CancellationToken(true)));
        Assert.Throws<ArgumentException>(() => review.ApplyResult(document, Array.Empty<PdfPageOrientationProposal>()));
        Assert.Equal(before, document.ToBytes());
    }

    [Fact]
    public async Task WholePageAndResourceLimitsRunBeforeProviderAndUnsupportedIsEvidence() {
        int calls = 0;
        var engine = new DelegateOcrEngine("unavailable", (_, _) => { calls++; return Task.FromResult(new OcrResult()); });
        var source = Source();
        var review = await source.PrepareOrientationReviewAsync(engine);
        Assert.Equal(0, calls); Assert.All(review.Pages, page => Assert.False(page.CanApply));
        await Assert.ThrowsAsync<ArgumentException>(() => source.PrepareOrientationReviewAsync(engine,
            new() { Regions = new[] { new PdfOcrPageRegion(1, 0, 0, 1, 1) } }));
        await Assert.ThrowsAsync<PdfReadLimitException>(() => source.PrepareOrientationReviewAsync(engine, new() { MaxPages = 1 }));
        await Assert.ThrowsAsync<PdfReadLimitException>(() => source.PrepareOrientationReviewAsync(Engine(90, 1), new() { MaxPixelsPerPage = 1 }));
        Assert.Equal(0, calls);
    }

    private static IOcrEngine Engine(int degrees, double confidence) => new DelegateOcrEngine("orientation-review", (request, _) => {
        Assert.Equal(OcrOperation.DetectOrientation, request.Operation);
        return Task.FromResult(new OcrResult { Provider = "fixture", Orientation = new() { ClockwiseRotationDegrees = degrees, Confidence = confidence } });
    }, new OcrEngineCapabilities { SupportsOrientationDetection = true });
    private static PdfDocument Source() => PdfDocument.Create(compose => {
        for (int i = 0; i < 2; i++) compose.Page(page => page.Size(300, 500).Content(content => content.Text("Orientation review")));
    });
}
