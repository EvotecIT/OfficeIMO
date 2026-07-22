using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOcrTests {
    [Fact]
    public async Task RecognizeAndMergeAsync_NormalizesFiltersAndMergesProviderWords() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Native text"))
            .ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            new PdfOcrWord("Native", 150, 140, 100, 30, 0.99),
            new PdfOcrWord("Scanned", 150, 400, 120, 32, 0.95),
            new PdfOcrWord("Weak", 300, 400, 80, 30, 0.2),
            new PdfOcrWord("Outside", request.PixelWidth, 0, 20, 20, 0.99)
        }, new[] { "provider-proof" }));

        PdfOcrMergeResult result = await PdfDocument.Open(pdf).Read.OcrAsync(provider);
        PdfOcrPageMergeResult page = Assert.Single(result.Pages);
        PdfRecognizedWord word = Assert.Single(page.Words);

        Assert.Equal(1, provider.CallCount);
        Assert.Equal(1, provider.LastRequest!.PageNumber);
        Assert.True(provider.LastRequest.Png.Length > 8);
        Assert.Equal("Scanned", word.Text);
        Assert.InRange(word.Confidence, 0.94, 0.96);
        Assert.Equal(1, page.RejectedLowConfidenceCount);
        Assert.Equal(1, page.RejectedNativeOverlapCount);
        Assert.Contains("provider-proof", page.Diagnostics);
        Assert.Contains(page.Diagnostics, diagnostic => diagnostic.StartsWith("InvalidWordGeometry:", StringComparison.Ordinal));
        Assert.Contains("Native text", page.Text, StringComparison.Ordinal);
        Assert.Contains("Scanned", page.Text, StringComparison.Ordinal);
        Assert.Same(result.NativeDocument.Pages[0], Assert.Single(result.NativeDocument.Pages));
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_HonorsSelectionAndCancellation() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("One"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Two"))
            .ToBytes();
        var provider = new StubOcrProvider(_ => new PdfOcrResponse(Array.Empty<PdfOcrWord>()));

        PdfOcrMergeResult selected = await PdfOcr.RecognizeAndMergeAsync(pdf, provider, new PdfOcrMergeOptions {
            Selection = PdfPageSelection.From(2)
        });
        Assert.Equal(2, Assert.Single(selected.Pages).PageNumber);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            PdfOcr.RecognizeAndMergeAsync(pdf, provider, cancellationToken: cancellation.Token));
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_RejectsOversizedProviderArtifactsBeforeMerge() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Native")).ToBytes();
        var provider = new StubOcrProvider(_ => new PdfOcrResponse(new[] {
            new PdfOcrWord("one", 10, 10, 10, 10, 0.9),
            new PdfOcrWord("two", 30, 10, 10, 10, 0.9)
        }));

        PdfReadLimitException exception = await Assert.ThrowsAsync<PdfReadLimitException>(() =>
            PdfOcr.RecognizeAndMergeAsync(pdf, provider, new PdfOcrMergeOptions {
                MaxOcrWordsPerPage = 1
            }));

        Assert.Equal(PdfReadLimitKind.OcrArtifacts, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    private sealed class StubOcrProvider : IPdfOcrProvider {
        private readonly Func<PdfOcrRequest, PdfOcrResponse> _response;
        public StubOcrProvider(Func<PdfOcrRequest, PdfOcrResponse> response) { _response = response; }
        public int CallCount { get; private set; }
        public PdfOcrRequest? LastRequest { get; private set; }
        public Task<PdfOcrResponse> RecognizeAsync(PdfOcrRequest request, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            CallCount++;
            LastRequest = request;
            return Task.FromResult(_response(request));
        }
    }
}
