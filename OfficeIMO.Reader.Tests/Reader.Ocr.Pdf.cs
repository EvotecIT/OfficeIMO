using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrPdfTests {
    [Theory]
    [InlineData(OfficeOcrCoordinateUnit.Pixels, 20D, 30D, 40D, 10D)]
    [InlineData(OfficeOcrCoordinateUnit.Points, 40D, 60D, 80D, 20D)]
    [InlineData(OfficeOcrCoordinateUnit.Normalized, 200D, 150D, 400D, 50D)]
    public async Task OfficeOcrEnginePdfProvider_ProjectsGeometryAndProvenance(
        OfficeOcrCoordinateUnit coordinateUnit,
        double expectedX,
        double expectedY,
        double expectedWidth,
        double expectedHeight) {
        var engine = new StubEngine(new OfficeOcrEngineResult {
            Provider = "fixture-provider",
            Model = "fixture-model",
            Language = "pol",
            Spans = new[] {
                new OfficeOcrTextSpan {
                    Sequence = 0,
                    Level = OfficeOcrTextSpanLevel.Word,
                    Text = "Invoice",
                    Confidence = 0.92D,
                    Region = coordinateUnit == OfficeOcrCoordinateUnit.Normalized
                        ? new OfficeDocumentRegion { X = 0.1D, Y = 0.1D, Width = 0.2D, Height = 1D / 30D }
                        : new OfficeDocumentRegion { X = 20D, Y = 30D, Width = 40D, Height = 10D },
                    CoordinateUnit = coordinateUnit
                }
            }
        });
        var provider = new OfficeOcrEnginePdfProvider(engine, new OfficeOcrEnginePdfProviderOptions {
            Language = "eng+pol",
            SourceName = "scan.pdf",
            SourceId = "scan-1"
        });

        PdfOcrResponse response = await provider.RecognizeAsync(new PdfOcrRequest(1, new byte[] { 1, 2, 3 }, 2000, 1500, 1000, 750, 2D));

        PdfOcrWord word = Assert.Single(response.Words);
        Assert.Equal(expectedX, word.X, 6);
        Assert.Equal(expectedY, word.Y, 6);
        Assert.Equal(expectedWidth, word.Width, 6);
        Assert.Equal(expectedHeight, word.Height, 6);
        Assert.Equal(0.92D, word.Confidence, 6);
        Assert.Equal("fixture-provider", response.Provider);
        Assert.Equal("fixture-model", response.Model);
        Assert.Equal("pol", response.Language);
        Assert.Equal("eng+pol", engine.LastRequest!.Language);
        Assert.Equal("scan.pdf", engine.LastRequest.Source.Path);
        Assert.Equal(1, engine.LastRequest.Candidate.Location.Page);
        Assert.Equal(2000, engine.LastRequest.Asset.Width);
    }

    [Fact]
    public async Task OfficeOcrEnginePdfProvider_UsesLineGeometryAndReportsMissingConfidence() {
        var engine = new StubEngine(new OfficeOcrEngineResult {
            Text = "Line result",
            Spans = new[] {
                new OfficeOcrTextSpan {
                    Sequence = 0,
                    Level = OfficeOcrTextSpanLevel.Line,
                    Text = "Line result",
                    Region = new OfficeDocumentRegion { X = 10D, Y = 20D, Width = 80D, Height = 12D }
                }
            }
        });
        var provider = new OfficeOcrEnginePdfProvider(engine, new OfficeOcrEnginePdfProviderOptions {
            ConfidenceWhenUnavailable = 0.75D
        });

        PdfOcrResponse response = await provider.RecognizeAsync(new PdfOcrRequest(1, new byte[] { 1 }, 100, 100, 50, 50, 2D));

        Assert.Equal("Line result", Assert.Single(response.Words).Text);
        Assert.Equal(0.75D, response.Words[0].Confidence, 6);
        Assert.Contains(response.Diagnostics, diagnostic => diagnostic.StartsWith("ocr-confidence-unavailable:", StringComparison.Ordinal));
    }

    [Fact]
    public async Task OfficeOcrEnginePdfProvider_DoesNotInventGeometryForPlainText() {
        var provider = new OfficeOcrEnginePdfProvider(new StubEngine(new OfficeOcrEngineResult { Text = "Text only" }));

        PdfOcrResponse response = await provider.RecognizeAsync(new PdfOcrRequest(1, new byte[] { 1 }, 100, 100, 50, 50, 2D));

        Assert.Empty(response.Words);
        Assert.Contains(response.Diagnostics, diagnostic => diagnostic.StartsWith("ocr-span-geometry-missing:", StringComparison.Ordinal));
    }

    [Fact]
    public async Task OfficeOcrEnginePdfProvider_IsolatesMutableEnginePayloadFromProvenanceEvidence() {
        byte[] callerPayload = { 1, 2, 3 };
        OfficeOcrEngineRequest? captured = null;
        var engine = new DelegateOfficeOcrEngine("mutating-fixture", (request, _) => {
            captured = request;
            request.Payload[0] = 99;
            return new ValueTask<OfficeOcrEngineResult>(new OfficeOcrEngineResult());
        });
        var provider = new OfficeOcrEnginePdfProvider(engine);

        await provider.RecognizeAsync(new PdfOcrRequest(1, callerPayload, 100, 100, 50, 50, 2D));

        Assert.NotNull(captured);
        Assert.NotSame(captured!.Payload, captured.Asset.PayloadBytes);
        Assert.Equal(new byte[] { 1, 2, 3 }, callerPayload);
        Assert.Equal(new byte[] { 1, 2, 3 }, captured.Asset.PayloadBytes);
        Assert.True(captured.Asset.PayloadHashMatches(out _));
    }

    private sealed class StubEngine : IOfficeOcrEngine {
        private readonly OfficeOcrEngineResult _result;

        internal StubEngine(OfficeOcrEngineResult result) {
            _result = result;
        }

        public string Id => "fixture-engine";

        public OfficeOcrEngineCapabilities Capabilities { get; } = new OfficeOcrEngineCapabilities {
            SupportsLineSpans = true,
            SupportsWordSpans = true,
            SupportsConfidence = true
        };

        internal OfficeOcrEngineRequest? LastRequest { get; private set; }

        public ValueTask<OfficeOcrEngineResult> RecognizeAsync(OfficeOcrEngineRequest request, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            LastRequest = request;
            return new ValueTask<OfficeOcrEngineResult>(_result);
        }
    }
}
