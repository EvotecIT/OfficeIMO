using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOcrRedactionSearchTests {
    [Fact]
    public async Task MultiWordLiteralMapsToOnePrivacySafeUserSpaceCandidate() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var engine = new DelegateOcrEngine(
            "fixture-ocr",
            (request, _) => Task.FromResult(new OcrResult {
                Provider = "fixture-provider",
                Model = "fixture-model",
                Language = "en",
                Spans = new[] {
                    Word("Account", 20, 30, 45, 12, 0.96),
                    Word("Secret", 70, 30, 38, 12, 0.91),
                    Word("Visible", 20, 60, 40, 12, 0.99)
                }
            }),
            new OcrEngineCapabilities { SupportsWordSpans = true, SupportsConfidence = true });
        var search = new PdfRedactionSearchOptions().AddLiteral("Account Secret");

        PdfOcrRedactionSearchResult result = await PdfDocument.Load(source)
            .SearchRedactionCandidatesWithOcrAsync(engine, search);

        PdfOcrRedactionCandidate candidate = Assert.Single(result.Candidates);
        Assert.Equal("literal:0", candidate.Criterion);
        Assert.Equal(0.91, candidate.MinimumConfidence, 2);
        Assert.Equal("fixture-provider", candidate.Provider);
        Assert.True(candidate.Area.Width > 80);
        Assert.DoesNotContain("Account Secret", candidate.Criterion, StringComparison.Ordinal);
    }

    [Fact]
    public async Task CancellationStopsBeforeProviderWork() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("page")).ToBytes();
        bool invoked = false;
        var engine = new DelegateOcrEngine("never", (_, _) => { invoked = true; return Task.FromResult(new OcrResult()); });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        var search = new PdfRedactionSearchOptions { CancellationToken = cancellation.Token }.AddLiteral("secret");
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => PdfDocument.Load(source)
            .SearchRedactionCandidatesWithOcrAsync(engine, search));

        Assert.False(invoked);
    }

    [Fact]
    public async Task LiteralDoesNotJoinWordsAcrossProviderLines() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("page")).ToBytes();
        var engine = new DelegateOcrEngine("fixture-ocr", (_, _) => Task.FromResult(new OcrResult {
            Spans = new[] {
                Word("Account", 20, 30, 45, 12, 0.96, "line-1"),
                Word("Secret", 20, 60, 38, 12, 0.91, "line-2")
            }
        }));

        PdfOcrRedactionSearchResult result = await PdfDocument.Load(source)
            .SearchRedactionCandidatesWithOcrAsync(engine, new PdfRedactionSearchOptions().AddLiteral("Account Secret"));

        Assert.Empty(result.Candidates);
    }

    private static OcrTextSpan Word(string text, double x, double y, double width, double height, double confidence, string? lineId = null) => new() {
        Level = OcrTextSpanLevel.Word,
        Text = text,
        Confidence = confidence,
        LineId = lineId,
        CoordinateUnit = OcrCoordinateUnit.Points,
        Region = new OcrRegion { X = x, Y = y, Width = width, Height = height }
    };
}
