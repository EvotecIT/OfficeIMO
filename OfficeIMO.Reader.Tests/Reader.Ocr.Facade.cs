using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Ocr;
using OfficeIMO.Reader.Ocr.Tesseract;
using OfficeIMO.Tests.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrFacadeTests {
    [Fact]
    public async Task Session_RecognizesImageWithStableSourceEvidenceAndConfiguredLanguage() {
        OfficeOcrEngineRequest? captured = null;
        var engine = new DelegateOfficeOcrEngine("fixture", (request, _) => {
            captured = request;
            return new ValueTask<OfficeOcrEngineResult>(new OfficeOcrEngineResult {
                Text = "Recognized",
                Provider = "fixture"
            });
        });
        OfficeOcrSession session = CreateSession(engine, "eng+pol", new PdfOcrMergeOptions());

        OfficeOcrEngineResult result = await session.RecognizeImageAsync(new byte[] { 1, 2, 3 }, "image/png", "scan.png");

        Assert.Equal("Recognized", result.Text);
        Assert.NotNull(captured);
        Assert.Equal("eng+pol", captured!.Language);
        Assert.Equal("image/png", captured.Asset.MediaType);
        Assert.Equal(".png", captured.Asset.Extension);
        Assert.Equal("scan.png", captured.Asset.FileName);
        Assert.Equal(64, captured.Source.SourceHash!.Length);
        Assert.Equal(captured.Source.SourceHash, captured.Asset.PayloadHash);
        Assert.NotSame(captured.Payload, captured.Asset.PayloadBytes);
    }

    [Fact]
    public async Task Session_SnapshotsPdfPolicyAndUsesEngineNeutralWordGeometry() {
        var policy = new PdfOcrMergeOptions {
            MinimumConfidence = 0.5D,
            DetectAlignedTables = false
        };
        var engine = new DelegateOfficeOcrEngine(
            "fixture",
            (_, _) => new ValueTask<OfficeOcrEngineResult>(new OfficeOcrEngineResult {
                Text = "Searchable",
                Confidence = 0.8D,
                Provider = "fixture",
                Language = "eng",
                Spans = new[] {
                    new OfficeOcrTextSpan {
                        Sequence = 0,
                        Level = OfficeOcrTextSpanLevel.Word,
                        Text = "Searchable",
                        Confidence = 0.8D,
                        CoordinateUnit = OfficeOcrCoordinateUnit.Normalized,
                        Region = new OfficeDocumentRegion { X = 0.1D, Y = 0.1D, Width = 0.3D, Height = 0.05D }
                    }
                }
            }),
            new OfficeOcrEngineCapabilities {
                SupportedMediaTypes = new[] { "image/png" },
                SupportsWordSpans = true,
                SupportsConfidence = true
            });
        OfficeOcrSession session = CreateSession(engine, "eng", policy);
        policy.MinimumConfidence = 0.9D;
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();

        PdfSearchableOcrResult result = await session.MakePdfSearchableAsync(PdfDocument.Load(source));

        Assert.Equal(1, result.AddedWordCount);
        Assert.Contains("Searchable", PdfReadDocument.Open(result.Document.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    private static OfficeOcrSession CreateSession(
        IOfficeOcrEngine engine,
        string language,
        PdfOcrMergeOptions options) {
        var runtime = new TesseractRuntimeInfo("fixture-tesseract", null, TesseractRuntimeSource.Explicit);
        var evidence = new OfficeOcrRuntimeEvidence(runtime, "fixture-version", new[] { "eng", "pol" }, null);
        return new OfficeOcrSession(engine, language, options, evidence);
    }
}
