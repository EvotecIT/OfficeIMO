using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Reader;
using OfficeIMO.Tests.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOcrPdfTests {
    [Fact]
    public async Task SharedEngineContract_IsReusableAcrossReaderAndPdfIntegrations() {
        var candidateKinds = new List<string?>();
        var engine = new DelegateOcrEngine(
            "shared-fixture",
            (request, _) => {
                candidateKinds.Add(request.CandidateKind);
                return Task.FromResult(request.CandidateKind == "page"
                    ? new OcrResult {
                        Text = "PDF scan",
                        Spans = new[] {
                            new OcrTextSpan {
                                Level = OcrTextSpanLevel.Word,
                                Text = "PDF scan",
                                Confidence = 0.95D,
                                CoordinateUnit = OcrCoordinateUnit.Normalized,
                                Region = new OcrRegion { X = 0.1D, Y = 0.2D, Width = 0.2D, Height = 0.04D }
                            }
                        }
                    }
                    : new OcrResult { Text = "Office image" });
            },
            new OcrEngineCapabilities {
                SupportedMediaTypes = new[] { "image/png" },
                SupportsWordSpans = true,
                SupportsConfidence = true,
                SupportsConcurrentRequests = true
            });

        OfficeDocumentOcrExecutionResult readerResult = await CreateReaderImageCandidate().ApplyOcrAsync(engine);
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), 220, 90)
            .ToBytes();
        PdfOcrMergeResult pdfResult = await PdfDocument.Load(pdf).ReadWithOcrAsync(engine);

        Assert.Equal("Office image", Assert.Single(readerResult.Recognitions).Result.Text);
        Assert.Equal("PDF scan", Assert.Single(pdfResult.Pages).Words[0].Text);
        Assert.Equal(new[] { "image", "page" }, candidateKinds);
    }

    [Fact]
    public async Task SharedRunner_SerializesOneNonConcurrentEngineAcrossReaderAndPdf() {
        var engine = new CrossIntegrationSerialOcrEngine();
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), 220, 90)
            .ToBytes();

        try {
            Task<OfficeDocumentOcrExecutionResult> reader = CreateReaderImageCandidate().ApplyOcrAsync(engine);
            Assert.Same(engine.FirstCallStarted, await Task.WhenAny(engine.FirstCallStarted, Task.Delay(TimeSpan.FromSeconds(10))));
            Task<PdfOcrMergeResult> pdfRead = PdfDocument.Load(pdf).ReadWithOcrAsync(engine);

            Assert.NotSame(engine.SecondCallStarted, await Task.WhenAny(engine.SecondCallStarted, Task.Delay(TimeSpan.FromMilliseconds(100))));
            engine.CompleteFirstCall();
            await Task.WhenAll(reader, pdfRead);

            Assert.Equal(1, engine.MaximumConcurrentCalls);
            Assert.Equal(2, engine.CallCount);
        } finally {
            engine.CompleteFirstCall();
        }
    }

    [Theory]
    [InlineData(OcrCoordinateUnit.Pixels)]
    [InlineData(OcrCoordinateUnit.Points)]
    [InlineData(OcrCoordinateUnit.Normalized)]
    public async Task PdfIntegration_ProjectsEveryNeutralCoordinateUnit(OcrCoordinateUnit coordinateUnit) {
        var engine = new DelegateOcrEngine(
            "geometry-fixture",
            (request, _) => Task.FromResult(new OcrResult {
                Provider = "fixture-provider",
                Model = "fixture-model",
                Language = "swe",
                Spans = new[] { CreateSpan(request, coordinateUnit) }
            }),
            new OcrEngineCapabilities {
                SupportedMediaTypes = new[] { "image/png" },
                SupportsWordSpans = true,
                SupportsConfidence = true
            });
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), 220, 90)
            .ToBytes();

        PdfOcrPageMergeResult page = Assert.Single((await PdfDocument.Load(pdf).ReadWithOcrAsync(engine)).Pages);
        PdfRecognizedWord word = Assert.Single(page.Words);

        Assert.Equal(20D, word.X, 3);
        Assert.Equal(30D, word.Y, 3);
        Assert.Equal(40D, word.Width, 3);
        Assert.Equal(10D, word.Height, 3);
        Assert.Equal("fixture-provider", page.Provider);
        Assert.Equal("fixture-model", page.Model);
        Assert.Equal("swe", page.Language);
    }

    private static OcrTextSpan CreateSpan(OcrRequest request, OcrCoordinateUnit unit) {
        OcrRegion page = request.Region ?? throw new InvalidOperationException("PDF page geometry is required.");
        double scaleX = request.PixelWidth.GetValueOrDefault() / page.Width;
        double scaleY = request.PixelHeight.GetValueOrDefault() / page.Height;
        OcrRegion region = unit switch {
            OcrCoordinateUnit.Pixels => new OcrRegion { X = 20D * scaleX, Y = 30D * scaleY, Width = 40D * scaleX, Height = 10D * scaleY },
            OcrCoordinateUnit.Points => new OcrRegion { X = 20D, Y = 30D, Width = 40D, Height = 10D },
            OcrCoordinateUnit.Normalized => new OcrRegion { X = 20D / page.Width, Y = 30D / page.Height, Width = 40D / page.Width, Height = 10D / page.Height },
            _ => throw new ArgumentOutOfRangeException(nameof(unit))
        };
        return new OcrTextSpan {
            Level = OcrTextSpanLevel.Word,
            Text = "Invoice",
            Confidence = 0.92D,
            BlockId = "block-2",
            ParagraphId = "paragraph-3",
            LineId = "line-4",
            Region = region,
            CoordinateUnit = unit
        };
    }

    private static OfficeDocumentReadResult CreateReaderImageCandidate() {
        byte[] payload = { 1, 2, 3 };
        var location = new ReaderLocation { Path = "workbook.xlsx", Page = 1, SourceBlockKind = "image", BlockAnchor = "asset-1" };
        var candidate = new OfficeDocumentOcrCandidate {
            Id = "ocr-1",
            Kind = "image",
            AssetId = "asset-1",
            Location = location,
            Region = new OfficeDocumentRegion { X = 0D, Y = 0D, Width = 10D, Height = 10D }
        };
        return new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Excel,
            Source = new OfficeDocumentSource { Path = "workbook.xlsx", SourceId = "workbook-1" },
            Assets = new[] {
                new OfficeDocumentAsset {
                    Id = "asset-1",
                    Kind = "image",
                    MediaType = "image/png",
                    Extension = ".png",
                    LengthBytes = payload.LongLength,
                    PayloadBytes = payload,
                    PayloadHash = OfficeDocumentAssetHash.ComputeSha256Hex(payload),
                    Location = location
                }
            },
            OcrCandidates = new[] { candidate },
            Pages = new[] { new OfficeDocumentPage { Number = 1, Location = location, OcrCandidates = new[] { candidate } } }
        };
    }

    private sealed class CrossIntegrationSerialOcrEngine : IOcrEngine {
        private readonly TaskCompletionSource<object?> _completeFirstCall =
            new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _firstCallStarted =
            new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<object?> _secondCallStarted =
            new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _activeCalls;
        private int _callCount;
        private int _maximumConcurrentCalls;

        public string Id => "cross-integration-serial";

        public OcrEngineCapabilities Capabilities { get; } = new OcrEngineCapabilities {
            SupportedMediaTypes = new[] { "image/png" },
            SupportsWordSpans = true,
            SupportsConcurrentRequests = false
        };

        internal Task FirstCallStarted => _firstCallStarted.Task;
        internal Task SecondCallStarted => _secondCallStarted.Task;
        internal int CallCount => Volatile.Read(ref _callCount);
        internal int MaximumConcurrentCalls => Volatile.Read(ref _maximumConcurrentCalls);

        public async Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
            int call = Interlocked.Increment(ref _callCount);
            int active = Interlocked.Increment(ref _activeCalls);
            UpdateMaximum(active);
            if (call == 1) _firstCallStarted.TrySetResult(null);
            if (call == 2) _secondCallStarted.TrySetResult(null);
            try {
                if (call == 1) await _completeFirstCall.Task.ConfigureAwait(false);
                if (request.CandidateKind == "page") {
                    return new OcrResult {
                        Text = "PDF scan",
                        Spans = new[] {
                            new OcrTextSpan {
                                Level = OcrTextSpanLevel.Word,
                                Text = "PDF scan",
                                Confidence = 0.95D,
                                CoordinateUnit = OcrCoordinateUnit.Points,
                                Region = new OcrRegion { X = 20D, Y = 30D, Width = 40D, Height = 10D }
                            }
                        }
                    };
                }
                return new OcrResult { Text = "Office image" };
            } finally {
                Interlocked.Decrement(ref _activeCalls);
            }
        }

        internal void CompleteFirstCall() {
            _completeFirstCall.TrySetResult(null);
        }

        private void UpdateMaximum(int active) {
            while (true) {
                int current = Volatile.Read(ref _maximumConcurrentCalls);
                if (active <= current || Interlocked.CompareExchange(ref _maximumConcurrentCalls, active, current) == current) return;
            }
        }
    }
}
