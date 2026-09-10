using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOcrScanProcessingTests {
    [Theory]
    [InlineData(72, 0)]
    [InlineData(144, 0)]
    [InlineData(144, 1)]
    public async Task RegionRecognitionSendsCroppedPixelsAndMapsWordsToOriginalPage(double dpi, int turns) {
        byte[] source = Source();
        var engine = new DelegateOcrEngine("region", (request, providerToken) => {
            int size = (int)(50 * dpi / 72);
            Assert.Equal(size, request.PixelWidth); Assert.Equal(size, request.PixelHeight);
            Assert.True(OfficeRasterImageDecoder.TryDecode(request.Payload, new OfficeRasterDecodeOptions(), out var cropped, out _));
            Assert.True(cropped!.GetPixel(size / 2, size / 2).R > cropped.GetPixel(size / 2, size / 2).B);
            return Task.FromResult(new OcrResult {
                Spans = new[] { new OcrTextSpan {
                Text = "Region", Level = OcrTextSpanLevel.Word, Confidence = 1, CoordinateUnit = OcrCoordinateUnit.Normalized,
                Region = new OcrRegion { X = .2, Y = .2, Width = .2, Height = .2 }
            } }
            });
        });
        var review = await PdfDocument.Load(source).PrepareSearchableOcrAsync(engine, new PdfOcrMergeOptions {
            Dpi = dpi,
            Regions = new[] { new PdfOcrPageRegion(1, .125, .25, .25, .5) },
            ScanProcessing = new() { ClockwiseQuarterTurns = turns, Deskew = false, NormalizeBackground = false, ColorMode = OfficeScanColorMode.PreserveColor }
        });
        var word = Assert.Single(review.Ocr.Pages[0].Words);
        Assert.Equal(35, word.X, 5); Assert.Equal(turns == 0 ? 35 : 55, word.Y, 5);
        Assert.Equal(10, word.Width, 5); Assert.Equal(10, word.Height, 5);
        byte[] searchable = review.ApplyAll().Document.ToBytes();
        Assert.Equal(PdfPageImageRenderer.RenderPageAsPng(source), PdfPageImageRenderer.RenderPageAsPng(searchable));
        Assert.Contains("Region", PdfReadDocument.Open(searchable).ExtractText());
    }

    [Fact]
    public async Task PerspectiveRecognitionMapsWordCornersAndKeepsSourceAppearance() {
        byte[] source = Source();
        var options = new OfficeScanPerspectiveOptions { TopLeft = new(.2, .1), TopRight = new(.8, .1), BottomRight = new(.95, .9), BottomLeft = new(.05, .9) };
        var engine = new DelegateOcrEngine("perspective", (request, _) => Task.FromResult(new OcrResult {
            Spans = new[] {
            new OcrTextSpan { Text = "Corrected", Level = OcrTextSpanLevel.Word, Confidence = 1, CoordinateUnit = OcrCoordinateUnit.Normalized,
                Region = new OcrRegion { X = .1, Y = .1, Width = .3, Height = .1 } }
        }
        }));
        var review = await PdfDocument.Load(source).PrepareSearchableOcrAsync(engine, new PdfOcrMergeOptions { Dpi = 72, Perspective = options });
        var expected = OfficeScanProcessor.CorrectPerspective(new OfficeRasterImage(200, 100, OfficeColor.White), options);
        var corner = expected.Mapping.MapProcessedToSource(new OfficePoint(expected.Image.Width * .1, expected.Image.Height * .1));
        var word = Assert.Single(review.Ocr.Pages[0].Words);
        Assert.Equal(corner.X, word.Geometry.TopLeft.X, 5); Assert.Equal(corner.Y, word.Geometry.TopLeft.Y, 5);
        Assert.Equal(PdfPageImageRenderer.RenderPageAsPng(source), PdfPageImageRenderer.RenderPageAsPng(review.ApplyAll().Document.ToBytes()));
    }

    [Fact]
    public async Task InvalidRegionPageFailsBeforeProviderWork() {
        int calls = 0;
        var engine = new DelegateOcrEngine("regions", (_, _) => { calls++; return Task.FromResult(new OcrResult()); });
        await Assert.ThrowsAsync<ArgumentException>(() => PdfDocument.Load(Source()).ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { Regions = new[] { new PdfOcrPageRegion(2, 0, 0, 1, 1) } }));
        Assert.Equal(0, calls);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task TransformedHierarchyFreeParagraphsRetainColumnOrder(int turns) {
        OcrTextSpan Word(string text, double x, double y) => new OcrTextSpan {
            Text = text,
            Level = OcrTextSpanLevel.Word,
            Confidence = 1,
            CoordinateUnit = OcrCoordinateUnit.Normalized,
            Region = new OcrRegion { X = x, Y = y, Width = 0.08, Height = 0.015 }
        };
        var words = new List<OcrTextSpan>();
        foreach (double y in new[] { 0.1, 0.25 }) {
            foreach (double x in new[] { 0.1, 0.6 }) {
                string name = x < 0.5 ? (y < 0.2 ? "Alpha" : "Beta") : (y < 0.2 ? "Gamma" : "Delta");
                words.Add(Word(name, x, y)); words.Add(Word("project", x + 0.09, y)); words.Add(Word("overview", x + 0.18, y));
            }
        }
        var engine = new DelegateOcrEngine("columns", (_, _) => Task.FromResult(new OcrResult { Spans = words }));
        var result = await PdfDocument.Load(Source()).ReadWithOcrAsync(engine, new PdfOcrMergeOptions {
            Dpi = 72,
            ScanProcessing = new OfficeScanProcessingOptions {
                ClockwiseQuarterTurns = turns,
                Deskew = false,
                NormalizeBackground = false,
                ColorMode = OfficeScanColorMode.PreserveColor
            }
        });
        Assert.Empty(result.Document.Tables);
        Assert.Equal("Alpha project overview Beta project overview Gamma project overview Delta project overview",
            System.Text.RegularExpressions.Regex.Replace(result.Text, @"\s+", " ").Trim());
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    [InlineData(3, true)]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    public async Task RotatedWordGeometryPreservesLogicalWordSpacing(int turns, bool hierarchy) {
        OcrTextSpan Word(string text, double x) => new OcrTextSpan {
            Text = text,
            Level = OcrTextSpanLevel.Word,
            Confidence = 1,
            LineId = hierarchy ? "line" : null,
            CoordinateUnit = OcrCoordinateUnit.Normalized,
            Region = new OcrRegion { X = x, Y = 0.1, Width = 0.15, Height = 0.05 }
        };
        var engine = new DelegateOcrEngine("rotated-spacing", (_, _) => Task.FromResult(new OcrResult {
            Spans = new[] { Word("Hello", 0.1), Word("world", 0.3) }
        }));
        var review = await PdfDocument.Load(Source()).PrepareSearchableOcrAsync(engine, new PdfOcrMergeOptions {
            Dpi = 72,
            ScanProcessing = new OfficeScanProcessingOptions {
                ClockwiseQuarterTurns = turns,
                Deskew = false,
                NormalizeBackground = false,
                ColorMode = OfficeScanColorMode.PreserveColor
            }
        });
        Assert.Equal("Hello world", review.Ocr.Text.Trim());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task TransformedHierarchyFreeLinesRetainReadingOrderAndPhraseSearch(int turns) {
        OcrTextSpan Word(string text, double x, double y) => new OcrTextSpan {
            Text = text,
            Level = OcrTextSpanLevel.Word,
            Confidence = 1,
            CoordinateUnit = OcrCoordinateUnit.Normalized,
            Region = new OcrRegion { X = x, Y = y, Width = 0.15, Height = 0.05 }
        };
        // A provider may return unsorted geometry without hierarchy. Infer rows in its recognized coordinate frame.
        var engine = new DelegateOcrEngine("rotated-lines", (_, _) => Task.FromResult(new OcrResult {
            Spans = new[] { Word("line", 0.3, 0.3), Word("world", 0.3, 0.1), Word("Next", 0.1, 0.3), Word("Hello", 0.1, 0.1) }
        }));
        var options = new PdfOcrMergeOptions {
            Dpi = 72,
            ScanProcessing = new OfficeScanProcessingOptions {
                ClockwiseQuarterTurns = turns,
                Deskew = false,
                NormalizeBackground = false,
                ColorMode = OfficeScanColorMode.PreserveColor
            }
        };
        PdfDocument document = PdfDocument.Load(Source());
        var review = await document.PrepareSearchableOcrAsync(engine, options);
        Assert.Equal("Hello world Next line", System.Text.RegularExpressions.Regex.Replace(review.Ocr.Text, @"\s+", " ").Trim());
        var search = await document.SearchRedactionCandidatesWithOcrAsync(engine,
            new PdfRedactionSearchOptions().AddLiteral("Hello world").AddLiteral("world Next"), options);
        PdfOcrRedactionCandidate candidate = Assert.Single(search.Candidates);
        Assert.Equal("literal:0", candidate.Criterion);
        Assert.True(candidate.Area.Width > 0); Assert.True(candidate.Area.Height > 0);
        var written = review.ApplyCorrections(review.Ocr.Pages[0].Words.ToDictionary(word => word,
            word => word.Text == "Hello" ? "Correct" : word.Text));
        Assert.Equal(PdfPageImageRenderer.RenderPageAsPng(Source()), PdfPageImageRenderer.RenderPageAsPng(written.Document.ToBytes()));
    }

    [Fact]
    public async Task DeskewWritesTheWordBaselineAtTheOriginalScanAngle() {
        byte[] original = System.IO.File.ReadAllBytes(System.IO.Path.Combine(AppContext.BaseDirectory, "ScanQuality", "phototest-skew.pdf"));
        var engine = new DelegateOcrEngine("deskew-geometry", (request, _) => Task.FromResult(new OcrResult {
            Spans = new[] { new OcrTextSpan { Text = "Mapped", Level = OcrTextSpanLevel.Word, Confidence = 1,
                CoordinateUnit = OcrCoordinateUnit.Normalized,
                Region = new OcrRegion { X = 0.25, Y = 0.25, Width = 0.2, Height = 0.05 } } }
        }));
        var review = await PdfDocument.Load(original).PrepareSearchableOcrAsync(engine, new PdfOcrMergeOptions {
            Dpi = 300,
            ScanProcessing = new OfficeScanProcessingOptions { NormalizeBackground = false }
        });
        PdfRecognizedWord word = Assert.Single(review.Ocr.Pages[0].Words);
        Assert.InRange(review.Ocr.Pages[0].ScanProcessing!.AppliedDeskewDegrees, -3.2, -2.8);
        double dx = word.Geometry.BottomRight.X - word.Geometry.BottomLeft.X;
        double dy = word.Geometry.BottomRight.Y - word.Geometry.BottomLeft.Y;
        Assert.InRange(dy / dx, Math.Tan(2.8 * Math.PI / 180), Math.Tan(3.2 * Math.PI / 180));
        byte[] searchable = review.ApplyAll().Document.ToBytes();
        Assert.Equal(PdfPageImageRenderer.RenderPageAsPng(original), PdfPageImageRenderer.RenderPageAsPng(searchable));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(searchable);
        var letters = independent.GetPage(1).Letters;
        Assert.Equal(6, letters.Count);
        double baselineDx = letters[5].EndBaseLine.X - letters[0].StartBaseLine.X;
        double baselineDy = letters[5].EndBaseLine.Y - letters[0].StartBaseLine.Y;
        Assert.InRange(baselineDy / baselineDx, -Math.Tan(3.2 * Math.PI / 180), -Math.Tan(2.8 * Math.PI / 180));
        Assert.Equal("Mapped", PdfReadDocument.Open(searchable).ExtractText().Trim());
    }

    [Theory]
    [InlineData(72, null)]
    [InlineData(144, null)]
    [InlineData(144, 100)]
    public async Task CorrectedLayerMapsRotatedAndDownsampledWordsToOriginalPage(double dpi, int? maximumDimension) {
        byte[] original = Source();
        var engine = new DelegateOcrEngine("geometry-fixture", (request, _) => {
            Assert.Equal(request.PixelHeight, request.PixelWidth * 2);
            return Task.FromResult(new OcrResult { Spans = new[] { RelativeWord(request) } });
        });
        var source = PdfDocument.Load(original);
        var review = await source.PrepareSearchableOcrAsync(engine, new PdfOcrMergeOptions {
            Dpi = dpi,
            ScanProcessing = new OfficeScanProcessingOptions {
                ClockwiseQuarterTurns = 1,
                Deskew = false,
                NormalizeBackground = false,
                ColorMode = OfficeScanColorMode.PreserveColor,
                MaximumDimension = maximumDimension
            }
        });
        PdfRecognizedWord word = Assert.Single(review.Ocr.Pages[0].Words);
        Assert.Equal(20D, word.X, 6); Assert.Equal(60D, word.Y, 6);
        Assert.Equal(10D, word.Width, 6); Assert.Equal(30D, word.Height, 6);
        Assert.Equal(20D, word.Geometry.TopLeft.X, 6); Assert.Equal(90D, word.Geometry.TopLeft.Y, 6);
        Assert.NotNull(review.Ocr.Pages[0].ScanProcessing);
        var written = review.ApplyCorrections(new Dictionary<PdfRecognizedWord, string> { [word] = "Fixed" });
        Assert.Same(word.Geometry, Assert.Single(written.WrittenWords[1]).Geometry);
        Assert.Equal(original, source.ToBytes());
        Assert.Equal(PdfPageImageRenderer.RenderPageAsPng(original), PdfPageImageRenderer.RenderPageAsPng(written.Document.ToBytes()));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(written.Document.ToBytes());
        var letters = independent.GetPage(1).Letters;
        Assert.Equal(5, letters.Count);
        Assert.InRange(letters[0].StartBaseLine.X, 29.99D, 30.01D);
        Assert.InRange(letters[0].StartBaseLine.Y, 9.99D, 10.01D);
        Assert.InRange(letters[4].EndBaseLine.X, 29.99D, 30.01D);
        Assert.InRange(letters[4].EndBaseLine.Y, 39.99D, 40.01D);
        Assert.Equal("Fixed", PdfReadDocument.Open(written.Document.ToBytes()).ExtractText().Trim());
    }

    [Theory]
    [InlineData(true, 0.9D, 100, 200)]
    [InlineData(true, 0.1D, 200, 100)]
    [InlineData(false, 0.9D, 200, 100)]
    public async Task OrientationUsesProviderEvidenceAndRetainsSourceWhenUnavailable(bool supported, double confidence, int width, int height) {
        int detections = 0, recognitions = 0;
        var engine = new DelegateOcrEngine("orientation-fixture", (request, _) => {
            if (request.Operation == OcrOperation.DetectOrientation) {
                detections++;
                return Task.FromResult(new OcrResult {
                    Orientation = new OcrOrientationResult { ClockwiseRotationDegrees = 90, Confidence = confidence }
                });
            }
            recognitions++;
            Assert.Equal(width, request.PixelWidth); Assert.Equal(height, request.PixelHeight);
            return Task.FromResult(new OcrResult());
        }, new OcrEngineCapabilities { SupportsOrientationDetection = supported });
        var result = await PdfDocument.Load(Source()).ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { DetectOrientation = true, Dpi = 72 });
        Assert.Equal(supported ? 1 : 0, detections); Assert.Equal(1, recognitions);
        Assert.Contains(result.Pages[0].Diagnostics, message => message.StartsWith("ocr-orientation", StringComparison.Ordinal));
    }

    [Fact]
    public async Task CleanupBudgetFallsBackToOriginalAndInvalidOptionsFailBeforeProviderWork() {
        int calls = 0;
        var engine = new DelegateOcrEngine("budget-fixture", (request, _) => {
            calls++; Assert.Equal(200, request.PixelWidth); Assert.Equal(100, request.PixelHeight);
            return Task.FromResult(new OcrResult());
        });
        var source = PdfDocument.Load(Source());
        var result = await source.ReadWithOcrAsync(engine, new PdfOcrMergeOptions {
            Dpi = 72,
            ScanProcessing = new OfficeScanProcessingOptions { MaximumPixels = 10 }
        });
        Assert.Equal(1, calls);
        Assert.Contains(result.Pages[0].Diagnostics, message => message.StartsWith("ocr-scan-limit:", StringComparison.Ordinal));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => source.ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { DetectOrientation = true, ScanProcessing = new OfficeScanProcessingOptions { BackgroundRadius = 0 } }));
        Assert.Equal(1, calls);
    }

    [Fact]
    public async Task OrientationAndRecognitionShareTheNonConcurrentEngineGate() {
        var firstEntered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        int calls = 0;
        var engine = new DelegateOcrEngine("shared-gate", async (request, _) => {
            int call = Interlocked.Increment(ref calls);
            if (call == 1) { firstEntered.SetResult(true); await release.Task.ConfigureAwait(false); }
            return new OcrResult();
        }, new OcrEngineCapabilities { SupportsOrientationDetection = true, SupportsConcurrentRequests = false });
        Task<OcrResult> orientation = OcrEngineRunner.RecognizeAsync(engine,
            new OcrRequest { Operation = OcrOperation.DetectOrientation }, TimeSpan.FromSeconds(10));
        // Releasing the provider must not wait for xUnit's bounded synchronization
        // context while other PDF tests occupy its workers.
#pragma warning disable xUnit1030 // The bounded provider handshake deliberately bypasses the test scheduler.
        await firstEntered.Task.ConfigureAwait(false);
        Task<OcrResult> recognition;
        try {
            recognition = OcrEngineRunner.RecognizeAsync(engine, new OcrRequest(), TimeSpan.FromSeconds(10));
            Assert.Equal(1, Volatile.Read(ref calls));
        } finally { release.TrySetResult(true); }
        await Task.WhenAll(orientation, recognition).ConfigureAwait(false);
#pragma warning restore xUnit1030
        Assert.Equal(2, calls);
    }

    private static OcrTextSpan RelativeWord(OcrRequest request) => new OcrTextSpan {
        Text = "Word",
        Level = OcrTextSpanLevel.Word,
        Confidence = 0.99D,
        CoordinateUnit = OcrCoordinateUnit.Pixels,
        Region = new OcrRegion {
            X = request.PixelWidth!.Value * 0.1D,
            Y = request.PixelHeight!.Value * 0.1D,
            Width = request.PixelWidth.Value * 0.3D,
            Height = request.PixelHeight.Value * 0.05D
        }
    };

    // An independent two-color scan producer; the PDF writer under test does not create this source.
    private static byte[] Source() {
        const string content = "q 200 0 0 100 0 0 cm /Im1 Do Q";
        return Encoding.ASCII.GetBytes("%PDF-1.4\n1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n" +
            "2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj\n" +
            "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >> endobj\n" +
            "4 0 obj << /Length " + content.Length + " >> stream\n" + content + "\nendstream endobj\n" +
            "5 0 obj << /Type /XObject /Subtype /Image /Width 2 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode /Length 13 >> stream\nFF00000000FF>\nendstream endobj\n" +
            "trailer << /Root 1 0 R >>\n%%EOF\n");
    }
}