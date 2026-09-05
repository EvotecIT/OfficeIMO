using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOcrTests {
    [Fact]
    public async Task RecognizeAndMergeAsync_NormalizesFiltersAndMergesProviderWords() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Native text"))
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            Word("Native", 150, 140, 100, 30, 0.99),
            Word("Scanned", 150, 400, 120, 32, 0.95),
            Word("Weak", 300, 400, 80, 30, 0.2),
            Word("Outside", request.PixelWidth.GetValueOrDefault(), 0, 20, 20, 0.99)
        }, new[] { "provider-proof" }, provider: "fixture", model: "fixture-v1", language: "eng"));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfOcrPageMergeResult page = Assert.Single(result.Pages);
        PdfRecognizedWord word = Assert.Single(page.Words);

        Assert.Equal(1, provider.CallCount);
        Assert.Equal(1, provider.LastRequest!.PageNumber);
        Assert.True(provider.LastRequest.Payload.Length > 8);
        Assert.Equal("Scanned", word.Text);
        Assert.InRange(word.Confidence, 0.94, 0.96);
        Assert.Equal(1, page.RejectedLowConfidenceCount);
        Assert.Equal(1, page.RejectedNativeOverlapCount);
        Assert.Contains("provider-proof", page.Diagnostics);
        Assert.Equal("fixture", page.Provider);
        Assert.Equal("fixture-v1", page.Model);
        Assert.Equal("eng", page.Language);
        Assert.Contains(page.Diagnostics, diagnostic => diagnostic.StartsWith("ocr-span-geometry:", StringComparison.Ordinal));
        Assert.Contains("Native text", page.Text, StringComparison.Ordinal);
        Assert.Contains("Scanned", page.Text, StringComparison.Ordinal);
        Assert.Same(result.NativeDocument.Pages[0], Assert.Single(result.NativeDocument.Pages));
        PdfLogicalTextBlock ocrBlock = Assert.Single(result.Document.TextBlocks, block => block.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.Equal("Scanned", ocrBlock.Text);
        Assert.NotNull(ocrBlock.VisualBounds);
        Assert.InRange(ocrBlock.Confidence, 0.94D, 0.96D);
        Assert.True(result.HasAcceptedOcrContent);
        Assert.Equal(1, result.AcceptedWordCount);
    }

    [Fact]
    public async Task Document_MergesNativeAndOcrBlocksInVisualReadingOrder() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Native follows OCR"))
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "OCR first", 30, 5, 70, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfLogicalPage page = Assert.Single(result.Document.Pages);

        Assert.Equal("OCR first", page.TextBlocks[0].Text);
        Assert.Equal(
            page.TextBlocks.Select(static block => block.Text),
            page.Elements.OfType<PdfLogicalTextBlock>().Select(static block => block.Text));
        Assert.True(page.TextBlocks.ToList().FindIndex(static block =>
            block.Text.Contains("Native follows OCR", StringComparison.Ordinal)) > 0);
        Assert.True(result.Document.Text.IndexOf("OCR first", StringComparison.Ordinal) <
            result.Document.Text.IndexOf("Native follows OCR", StringComparison.Ordinal));
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_HonorsSelectionAndCancellation() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("One"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Two"))
            .ToBytes();
        var provider = new StubOcrEngine(_ => Result(Array.Empty<OcrTextSpan>()));

        PdfOcrMergeResult selected = await PdfOcr.RecognizeAndMergeAsync(pdf, provider, new PdfOcrMergeOptions {
            ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(2) }
        });
        Assert.Equal(2, Assert.Single(selected.Pages).PageNumber);
        PdfLogicalPage selectedNativePage = Assert.Single(selected.NativeDocument.Pages);
        Assert.Equal(2, selectedNativePage.PageNumber);
        Assert.Contains(selectedNativePage.TextBlocks, block => block.Text.Contains("Two", StringComparison.Ordinal));
        Assert.DoesNotContain(selectedNativePage.TextBlocks, block => block.Text.Contains("One", StringComparison.Ordinal));

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            PdfOcr.RecognizeAndMergeAsync(pdf, provider, cancellationToken: cancellation.Token));
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_BoundsProviderLifetime() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), 220, 90)
            .ToBytes();
        var completion = new TaskCompletionSource<OcrResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        var provider = new DelegateOcrEngine("non-cooperative-pdf-fixture", (_, _) => completion.Task);

        try {
            await Assert.ThrowsAsync<OcrEngineTimeoutException>(() => PdfDocument.Load(pdf).ReadWithOcrAsync(
                provider,
                new PdfOcrMergeOptions { ProviderTimeout = TimeSpan.FromMilliseconds(20) }));
        } finally {
            completion.TrySetResult(new OcrResult());
        }
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_RejectsOversizedEngineIdentifierBeforeCallingProvider() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), 220, 90)
            .ToBytes();
        var provider = new StubOcrEngine(
            _ => new OcrResult(),
            new string(' ', OcrEngineRunner.MaximumEngineIdCharacters) + "x");

        await Assert.ThrowsAsync<ArgumentException>(() => PdfDocument.Load(pdf).ReadWithOcrAsync(provider));

        Assert.Equal(0, provider.CallCount);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_CapturesEngineIdentityAndCapabilitiesOncePerDocumentOperation() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .PageBreak()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new ChangingIdentityOcrEngine();

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);

        Assert.Equal(1, provider.IdReadCount);
        Assert.Equal(1, provider.CapabilitiesReadCount);
        Assert.Equal(2, provider.CallCount);
        Assert.Equal(2, result.Pages.Count);
        Assert.All(result.Pages, page => Assert.Equal("snapshot-engine", page.Provider));
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_DoesNotRestartTimedOutNonConcurrentProvider() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(230, 230, 230), 220, 90)
            .ToBytes();
        var completion = new TaskCompletionSource<OcrResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        var firstCallStarted = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        int callCount = 0;
        var provider = new DelegateOcrEngine("stalled-non-concurrent-pdf-fixture", (_, _) => {
            Interlocked.Increment(ref callCount);
            firstCallStarted.TrySetResult(null);
            return completion.Task;
        });
        var options = new PdfOcrMergeOptions { ProviderTimeout = TimeSpan.FromMilliseconds(200) };

        try {
            Task<PdfOcrMergeResult> firstCall = PdfDocument.Load(pdf).ReadWithOcrAsync(provider, options);
            Task started = await Task.WhenAny(firstCallStarted.Task, Task.Delay(TimeSpan.FromSeconds(10)));
            Assert.Same(firstCallStarted.Task, started);
            OcrEngineTimeoutException first = await Assert.ThrowsAsync<OcrEngineTimeoutException>(
                () => firstCall);
            OcrEngineTimeoutException second = await Assert.ThrowsAsync<OcrEngineTimeoutException>(
                () => PdfDocument.Load(pdf).ReadWithOcrAsync(provider, options));

            Assert.True(first.ProviderCallStarted);
            Assert.False(second.ProviderCallStarted);
            Assert.Equal(1, Volatile.Read(ref callCount));
        } finally {
            completion.TrySetResult(new OcrResult());
        }
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_UsesCanonicalStructuredReadForNativeEvidence() {
        byte[] pdf = PdfDocument.Create()
            .TaggedPdfCatalogMarkers()
            .H1("Native structured heading")
            .ToBytes();
        var provider = new StubOcrEngine(_ => Result(Array.Empty<OcrTextSpan>()));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfUnderstandingPageResult analysis = Assert.Single(result.NativeDocument.Pages).Analysis;

        Assert.Equal(PdfReadProfile.Structured, result.NativeDocument.Profile);
        Assert.Contains(analysis.Elements, element =>
            element.Kind == PdfUnderstandingSemanticKind.Heading &&
            element.Region.Text.Contains("Native structured heading", StringComparison.Ordinal) &&
            element.Evidence.Any(static evidence => evidence.Code == "semantic.tagged-pdf-role"));
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_AppliesDocumentEvidenceOnceAfterOcrFusion() {
        byte[] pdf = PdfDocument.Create()
            .Header(header => header.AlignLeft().Text("Shared running header"))
            .Paragraph(paragraph => paragraph.Text("First page body."))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page body."))
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(
            request.PageNumber == 1
                ? new[] { At(request, "Scanned", 300, 400, 52, 12) }
                : Array.Empty<OcrTextSpan>()));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);

        Assert.Equal(2, result.Document.Pages.Count);
        Assert.All(result.Document.Pages, page => {
            PdfUnderstandingSemanticElement header = Assert.Single(
                page.Analysis.Elements,
                element => element.Kind == PdfUnderstandingSemanticKind.Header &&
                    element.Region.Text.Contains("Shared running header", StringComparison.Ordinal));
            Assert.Equal(
                1,
                header.Evidence.Count(static evidence => evidence.Code == "semantic.repeated-header"));
        });
        Assert.Contains(
            result.Document.Pages[0].TextBlocks,
            block => block.SourceKind == PdfLogicalContentSourceKind.Ocr && block.Text == "Scanned");
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_AppliesCallerReadProfileAndSemanticStageToOcrEvidence() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Semantic", 30, 90, 58, 12),
            At(request, "evidence", 94, 90, 52, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider, new PdfOcrMergeOptions {
            ReadOptions = new PdfReadOptions {
                Profile = PdfReadProfile.Fast,
                Pipeline = new PdfUnderstandingPipelineOptions {
                    SemanticClassification = new OcrHeadingClassificationStage()
                }
            }
        });

        Assert.Equal(PdfReadProfile.Fast, result.NativeDocument.Profile);
        Assert.Equal(PdfReadProfile.Fast, result.Document.Profile);
        Assert.Equal("Semantic evidence", Assert.Single(result.Document.Headings).Text);
        Assert.Equal(
            typeof(OcrHeadingClassificationStage),
            Assert.Single(result.Document.Pages[0].Analysis.Trace, static trace => trace.Stage == "semantic-classification").ProviderType);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_UsesSharedReadPipelineLimits() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("One"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Two"))
            .ToBytes();
        var provider = new StubOcrEngine(_ => Result(Array.Empty<OcrTextSpan>()));

        PdfReadLimitException exception = await Assert.ThrowsAsync<PdfReadLimitException>(() =>
            PdfOcr.RecognizeAndMergeAsync(pdf, provider, new PdfOcrMergeOptions {
                MaxPages = 10,
                ReadOptions = new PdfReadOptions {
                    Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 1 }
                }
            }));

        Assert.Equal(PdfReadLimitKind.Pages, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(0, provider.CallCount);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_PreservesDuplicateCallerOrderedPages() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Repeated", 30, 90, 50, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider, new PdfOcrMergeOptions {
            ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(1, 1) }
        });

        Assert.Equal(2, provider.CallCount);
        Assert.Equal(new[] { 1, 1 }, result.Pages.Select(static page => page.PageNumber));
        Assert.Equal(new[] { 1, 1 }, result.Document.Pages.Select(static page => page.PageNumber));
        Assert.All(result.Document.Pages, page =>
            Assert.Single(page.TextBlocks, block => block.SourceKind == PdfLogicalContentSourceKind.Ocr));
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_RejectsOversizedProviderArtifactsBeforeMerge() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Native")).ToBytes();
        var provider = new StubOcrEngine(_ => Result(new[] {
            Word("one", 10, 10, 10, 10, 0.9),
            Word("two", 30, 10, 10, 10, 0.9)
        }));

        PdfReadLimitException exception = await Assert.ThrowsAsync<PdfReadLimitException>(() =>
            PdfOcr.RecognizeAndMergeAsync(pdf, provider, new PdfOcrMergeOptions {
                MaxOcrWordsPerPage = 1
            }));

        Assert.Equal(PdfReadLimitKind.OcrArtifacts, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_BoundsRawProviderMetadataBeforeTrimming() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        string padded = new string(' ', 32) + "tail";
        var provider = new StubOcrEngine(request => Result(
            new[] { At(request, "text", 20D, 20D, 10D, 10D) },
            provider: padded));

        PdfReadLimitException exception = await Assert.ThrowsAsync<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf).ReadWithOcrAsync(provider, new PdfOcrMergeOptions {
                MaxProviderMetadataCharactersPerPage = 16
            }));

        Assert.Equal(PdfReadLimitKind.OcrArtifacts, exception.Kind);
        Assert.Equal(16, exception.Limit);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_BoundsNativeOverlapWork() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First native block"))
            .Paragraph(paragraph => paragraph.Text("Second native block"))
            .ToBytes();
        var provider = new StubOcrEngine(_ => Result(new[] {
            Word("scanned", 10, 10, 10, 10, 0.9)
        }));

        PdfReadLimitException exception = await Assert.ThrowsAsync<PdfReadLimitException>(() =>
            PdfOcr.RecognizeAndMergeAsync(pdf, provider, new PdfOcrMergeOptions {
                MaxNativeTextOverlapComparisonsPerPage = 1
            }));

        Assert.Equal(PdfReadLimitKind.OcrArtifacts, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public async Task MakeSearchableAsync_RejectsOcrOverVisibleArtifactTextWithoutExposingArtifactsInTheLogicalResult() {
        byte[] source = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Canvas(canvas => canvas.Artifact(artifact => artifact.Text("Footer", 50D, 700D, 100D, 20D)))
            .ToBytes();
        PdfPageInteractionMap map = PdfPageInteractionMap.Create(
            source,
            1,
            readOptions: new PdfLoadOptions { IncludeArtifactText = true });
        PdfSelectionQuad[] glyphs = map.TextRegions.Select(static region => region.Quad).ToArray();
        Assert.NotEmpty(glyphs);
        double left = glyphs.Min(static quad => quad.Left);
        double top = glyphs.Min(static quad => quad.Top);
        double right = glyphs.Max(static quad => quad.Right);
        double bottom = glyphs.Max(static quad => quad.Bottom);
        var provider = new StubOcrEngine(request => Result(new[] {
            Word(
                "Footer",
                left * Scale(request),
                top * Scale(request),
                (right - left) * Scale(request),
                (bottom - top) * Scale(request),
                0.99D)
        }));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).MakeSearchableAsync(provider);

        PdfOcrPageMergeResult page = Assert.Single(result.Ocr.Pages);
        Assert.Empty(page.Words);
        Assert.Equal(1, page.RejectedNativeOverlapCount);
        Assert.Empty(result.Ocr.NativeDocument.TextBlocks);
        Assert.False(result.WasModified);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_CombinesFragmentedNativeSpansForOverlapRejection() {
        byte[] source = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Canvas(canvas => canvas
                .Text("N", 50D, 100D, 12D, 20D, fontSize: 12D)
                .Text("a", 62D, 100D, 12D, 20D, fontSize: 12D)
                .Text("t", 74D, 100D, 12D, 20D, fontSize: 12D)
                .Text("i", 86D, 100D, 12D, 20D, fontSize: 12D)
                .Text("v", 98D, 100D, 12D, 20D, fontSize: 12D)
                .Text("e", 110D, 100D, 12D, 20D, fontSize: 12D))
            .ToBytes();
        PdfSelectionQuad[] glyphs = PdfPageInteractionMap.Create(source, 1).TextRegions
            .Select(static region => region.Quad)
            .ToArray();
        Assert.Equal(6, glyphs.Length);
        double left = glyphs.Min(static quad => quad.Left);
        double top = glyphs.Min(static quad => quad.Top);
        double right = glyphs.Max(static quad => quad.Right);
        double bottom = glyphs.Max(static quad => quad.Bottom);
        var provider = new StubOcrEngine(request => Result(new[] {
            Word("Native", left * Scale(request), top * Scale(request),
                (right - left) * Scale(request), (bottom - top) * Scale(request), 0.99D)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(source).ReadWithOcrAsync(provider);

        PdfOcrPageMergeResult page = Assert.Single(result.Pages);
        Assert.Empty(page.Words);
        Assert.Equal(1, page.RejectedNativeOverlapCount);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_DoesNotDoubleCountOverlappingNativeSpans() {
        byte[] source = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Canvas(canvas => canvas
                .Text("A", 50D, 100D, 12D, 20D, fontSize: 12D)
                .Text("A", 50D, 100D, 12D, 20D, fontSize: 12D))
            .ToBytes();
        PdfSelectionQuad glyph = PdfPageInteractionMap.Create(source, 1).TextRegions[0].Quad;
        var provider = new StubOcrEngine(request => Result(new[] {
            Word("Scanned", glyph.Left * Scale(request), glyph.Top * Scale(request),
                glyph.Width * 3D * Scale(request), glyph.Height * Scale(request), 0.99D)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(source).ReadWithOcrAsync(provider);

        PdfOcrPageMergeResult page = Assert.Single(result.Pages);
        Assert.Equal("Scanned", Assert.Single(page.Words).Text);
        Assert.Equal(0, page.RejectedNativeOverlapCount);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_DoesNotUseFullyClippedTextForOverlapRejection() {
        byte[] source = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Canvas(canvas => canvas.Clip(10D, 10D, 10D, 10D, clipped =>
                clipped.Text("Clipped", 100D, 100D, 80D, 20D, fontSize: 12D)))
            .ToBytes();
        PdfSelectionQuad[] glyphs = PdfPageInteractionMap.Create(source, 1).TextRegions
            .Select(static region => region.Quad)
            .ToArray();
        Assert.NotEmpty(glyphs);
        double left = glyphs.Min(static quad => quad.Left);
        double top = glyphs.Min(static quad => quad.Top);
        double right = glyphs.Max(static quad => quad.Right);
        double bottom = glyphs.Max(static quad => quad.Bottom);
        var provider = new StubOcrEngine(request => Result(new[] {
            Word("Scanned", left * Scale(request), top * Scale(request),
                (right - left) * Scale(request), (bottom - top) * Scale(request), 0.99D)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(source).ReadWithOcrAsync(provider);

        PdfOcrPageMergeResult page = Assert.Single(result.Pages);
        Assert.Equal("Scanned", Assert.Single(page.Words).Text);
        Assert.Equal(0, page.RejectedNativeOverlapCount);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_UsesPaintedAdvanceForActualTextOverlapBounds() {
        byte[] source = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Canvas(canvas => canvas.ActualText("A much longer replacement", logical =>
                logical.Text("X", 50D, 100D, 12D, 20D, fontSize: 12D)))
            .ToBytes();
        PdfReadPage readPage = PdfReadDocument.Open(source).Pages[0];
        PdfTextSpan nativeSpan = Assert.Single(
            readPage.GetInteractionTextSpans(),
            static span => span.Text == "A much longer replacement");
        PdfSelectionQuad firstLogicalGlyph = PdfPageInteractionMap.Create(source, 1).TextRegions[0].Quad;
        double ocrLeft = firstLogicalGlyph.Left + Math.Abs(nativeSpan.Advance) + 2D;
        var provider = new StubOcrEngine(request => Result(new[] {
            Word("Scanned", ocrLeft * Scale(request), firstLogicalGlyph.Top * Scale(request),
                24D * Scale(request), firstLogicalGlyph.Height * Scale(request), 0.99D)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(source).ReadWithOcrAsync(provider);

        PdfOcrPageMergeResult page = Assert.Single(result.Pages);
        Assert.Equal("Scanned", Assert.Single(page.Words).Text);
        Assert.Equal(0, page.RejectedNativeOverlapCount);
    }

    [Theory]
    [InlineData(90)]
    [InlineData(270)]
    public async Task RecognizeAndMergeAsync_UsesVisualCoordinatesForRotatedCroppedPages(int rotation) {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Native rotation"))
            .ToBytes();
        pdf = PdfPageEditor.SetCropBox(pdf, 0, 200, 595, 842);
        pdf = PdfPageEditor.RotatePages(pdf, rotation);
        PdfPageInteractionMap map = PdfPageInteractionMap.Create(pdf, 1);
        Assert.NotEmpty(map.TextRegions);
        PdfSelectionQuad nativeGlyph = map.TextRegions[0].Quad;
        var provider = new StubOcrEngine(request => Result(new[] {
            Word(
                "duplicate",
                nativeGlyph.Left * Scale(request),
                nativeGlyph.Top * Scale(request),
                nativeGlyph.Width * Scale(request),
                nativeGlyph.Height * Scale(request),
                0.99)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfOcrPageMergeResult page = Assert.Single(result.Pages);

        Assert.Equal(map.Width, provider.LastRequest!.Region!.Width, 3);
        Assert.Equal(map.Height, provider.LastRequest.Region.Height, 3);
        Assert.Empty(page.Words);
        Assert.Equal(1, page.RejectedNativeOverlapCount);
    }

    [Fact]
    public void PdfReadLimitKind_PreservesExistingInteractionRegionsValue() {
        Assert.Equal(20, (int)PdfReadLimitKind.InteractionRegions);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public async Task Document_PreservesOcrVisualGeometryAcrossPageRotation(int rotation) {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Native anchor")).ToBytes();
        pdf = PdfPageEditor.SetCropBox(pdf, 20, 40, 500, 700);
        pdf = PdfPageEditor.RotatePages(pdf, rotation);
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "OCR", 24, 180, 42, 12),
            At(request, "geometry", 72, 180, 64, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfLogicalTextBlock block = Assert.Single(result.Document.TextBlocks, candidate => candidate.SourceKind == PdfLogicalContentSourceKind.Ocr);
        PdfLogicalVisualBounds bounds = Assert.IsType<PdfLogicalVisualBounds>(block.VisualBounds);
        Assert.Equal(24D, bounds.Left, 3);
        Assert.Equal(180D, bounds.Top, 3);
        Assert.Equal("OCR geometry", block.Text);
        PdfLogicalPage logicalPage = result.Document.Pages[0];
        PdfLogicalReadingOrderItem entry = Assert.Single(
            PdfLogicalReadingOrderAnalysis.Analyze(logicalPage),
            candidate => candidate.Kind == PdfLogicalReadingOrderKind.Paragraph &&
                logicalPage.Paragraphs[candidate.SourceIndex].Text.Contains("OCR geometry", StringComparison.Ordinal));
        Assert.Equal(24D, entry.Left, 3);
        Assert.Equal(180D, entry.Top, 3);
    }

    [Fact]
    public async Task Document_InfersOnlyRepeatedAlignedOcrRowsAsEditableTable() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Native anchor")).ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Item", 30, 220, 34, 12), At(request, "Value", 150, 220, 38, 12),
            At(request, "Alpha", 30, 245, 38, 12), At(request, "10", 150, 245, 18, 12),
            At(request, "Beta", 30, 270, 32, 12), At(request, "20", 150, 270, 18, 12),
            At(request, "Narrative", 30, 320, 58, 12), At(request, "continues", 94, 320, 52, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfLogicalTable table = Assert.Single(result.Document.Tables, candidate => candidate.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.Equal("ocr-aligned-geometry", table.DetectionKind);
        Assert.Equal(PdfTableCoordinateSpace.VisualTopLeft, table.CoordinateSpace);
        PdfUnderstandingTableCandidate candidate = Assert.Single(
            result.Document.Pages[0].Analysis.TableCandidates,
            candidate => candidate.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.Equal(table.Rows.Select(static row => row.ToArray()), candidate.Rows.Select(static row => row.ToArray()));
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(new[] { "Alpha", "10" }, table.Rows[1]);
        Assert.DoesNotContain(table.Rows, row => row.Contains("Narrative"));

        PdfExcelTableImportResult excelResult = result.Document.ImportTablesToExcelDocumentResult();
        using (excelResult.Value) Assert.Single(excelResult.Report.Entries);
        PdfOdsConversionResult odsResult = result.Document.ToOdsDocumentResult();
        Assert.Single(odsResult.Value.Sheets);

        PdfPowerPointConversionResult presentationResult = result.Document
            .ToPowerPointPresentationResult(PdfToPowerPointOptions.CreateEditableContent());
        using (presentationResult.Value) {
            Assert.Single(presentationResult.Report.EditablePages);
            Assert.Equal(1, presentationResult.Report.EditablePages[0].TableCount);
        }

        string html = result.Document.ToHtml();
        Assert.Equal(1, html.Split(new[] { "Alpha" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void Document_DoesNotDuplicateNativeTableWithOverlappingOcrCandidate() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Item", "Value" },
                new[] { "Alpha", "10" },
                new[] { "Beta", "20" }
            })
            .ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(pdf).Pages);
        PdfUnderstandingTableCandidate native = Assert.Single(page.Analysis.TableCandidates);
        double left = native.Columns.Min(static column => column.From);
        double right = native.Columns.Max(static column => column.To);
        PdfVisualBounds visual = page.TransformBoundsToVisual(
            left,
            Math.Min(native.YBottom, native.YTop),
            right,
            Math.Max(native.YBottom, native.YTop));
        PdfUnderstandingTableCandidate ocr = PdfUnderstandingTableCandidate.FromOcr(
            "OcrAlignedColumns",
            visual.Top - 2D,
            visual.Bottom + 2D,
            new PdfLogicalVisualBounds(visual.Left - 2D, visual.Top - 2D, visual.Right + 2D, visual.Bottom + 2D),
            native.Columns.Select(column => (column.From, column.To)).ToArray(),
            native.Rows,
            0.99D,
            new[] { new PdfInferenceEvidence("test.ocr-duplicate", "Overlapping OCR candidate.", 1D) });

        IReadOnlyList<PdfUnderstandingTableCandidate> reconciled =
            PdfUnderstandingTableCandidateReconciler.Reconcile(page, new[] { native }, new[] { ocr });

        Assert.Same(native, Assert.Single(reconciled));
    }

    [Fact]
    public void Document_DoesNotReplaceGeometryWhenTaggedTableOwnsOnlySomeNativeRuns() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Item", "Value" },
                new[] { "Alpha", "10" },
                new[] { "Beta", "20" }
            })
            .ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(pdf).Pages);
        PdfUnderstandingTableCandidate geometry = Assert.Single(page.Analysis.TableCandidates);
        PdfUnderstandingTableCandidate incompleteTagged = new(
            "tagged-structure",
            geometry.YTop,
            geometry.YBottom,
            geometry.Columns,
            geometry.Rows,
            geometry.SourceLines.Take(1).ToArray(),
            0.99D,
            new[] {
                new PdfInferenceEvidence(
                    "table.tagged-structure",
                    "Incomplete tagged ownership used by this regression test.",
                    0.99D)
            });

        IReadOnlyList<PdfUnderstandingTableCandidate> reconciled =
            PdfUnderstandingTableCandidateReconciler.Reconcile(
                page,
                new[] { geometry },
                new[] { incompleteTagged });

        Assert.Same(geometry, Assert.Single(reconciled));
    }

    [Fact]
    public void Document_PreservesRicherOverlappingOcrTable() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Item", "Value" },
                new[] { "Alpha", "10" }
            })
            .ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(pdf).Pages);
        PdfUnderstandingTableCandidate native = Assert.Single(page.Analysis.TableCandidates);
        double left = native.Columns.Min(static column => column.From);
        double right = native.Columns.Max(static column => column.To);
        PdfVisualBounds visual = page.TransformBoundsToVisual(
            left,
            Math.Min(native.YBottom, native.YTop),
            right,
            Math.Max(native.YBottom, native.YTop));
        var richerRows = native.Rows
            .Concat(new IReadOnlyList<string>[] { new[] { "Beta", "20" }, new[] { "Gamma", "30" } })
            .ToArray();
        PdfUnderstandingTableCandidate ocr = PdfUnderstandingTableCandidate.FromOcr(
            "OcrAlignedColumns",
            visual.Top - 2D,
            visual.Bottom + 40D,
            new PdfLogicalVisualBounds(visual.Left - 2D, visual.Top - 2D, visual.Right + 2D, visual.Bottom + 40D),
            native.Columns.Select(column => (column.From, column.To)).ToArray(),
            richerRows,
            0.99D,
            new[] { new PdfInferenceEvidence("test.ocr-richer", "OCR candidate contains additional rows.", 1D) });

        PdfUnderstandingTableCandidate accepted = Assert.Single(
            PdfUnderstandingTableCandidateReconciler.Reconcile(page, new[] { native }, new[] { ocr }));

        Assert.Same(ocr, accepted);
        Assert.Equal(PdfLogicalContentSourceKind.Ocr, accepted.SourceKind);
        Assert.Equal(4, accepted.Rows.Count);
        Assert.Contains(accepted.Rows, row => row.SequenceEqual(new[] { "Gamma", "30" }));
    }

    [Fact]
    public void Document_DoesNotCollapseOverlappingTablesThatOnlyShareAHeaderRow() {
        byte[] pdf = PdfDocument.Create()
            .Table(new[] {
                new[] { "Item", "Value" },
                new[] { "Alpha", "10" }
            })
            .ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocumentReadResult.Load(pdf).Pages);
        PdfUnderstandingTableCandidate native = Assert.Single(page.Analysis.TableCandidates);
        double left = native.Columns.Min(static column => column.From);
        double right = native.Columns.Max(static column => column.To);
        PdfVisualBounds visual = page.TransformBoundsToVisual(
            left,
            Math.Min(native.YBottom, native.YTop),
            right,
            Math.Max(native.YBottom, native.YTop));
        PdfUnderstandingTableCandidate distinct = PdfUnderstandingTableCandidate.FromOcr(
            "OcrAlignedColumns",
            visual.Top,
            visual.Bottom,
            new PdfLogicalVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom),
            native.Columns.Select(column => (column.From, column.To)).ToArray(),
            new IReadOnlyList<string>[] {
                new[] { "Item", "Value" },
                new[] { "Beta", "20" }
            },
            0.99D,
            new[] { new PdfInferenceEvidence("test.ocr-distinct", "Distinct overlapping table.", 1D) });

        IReadOnlyList<PdfUnderstandingTableCandidate> reconciled =
            PdfUnderstandingTableCandidateReconciler.Reconcile(page, new[] { native }, new[] { distinct });

        Assert.Equal(2, reconciled.Count);
    }

    [Fact]
    public async Task Document_KeepsAlignedNarrativeColumnsOutOfTableInference() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Alpha", 30, 50, 30, 10), At(request, "project", 65, 50, 42, 10), At(request, "overview", 112, 50, 48, 10),
            At(request, "Delta", 300, 50, 30, 10), At(request, "project", 335, 50, 42, 10), At(request, "overview", 382, 50, 48, 10),
            At(request, "continues", 30, 66, 48, 10), At(request, "with", 83, 66, 24, 10), At(request, "details", 112, 66, 36, 10),
            At(request, "continues", 300, 66, 48, 10), At(request, "with", 353, 66, 24, 10), At(request, "details", 382, 66, 36, 10),
            At(request, "for", 30, 82, 18, 10), At(request, "each", 53, 82, 24, 10), At(request, "audience", 82, 82, 48, 10),
            At(request, "for", 300, 82, 18, 10), At(request, "each", 323, 82, 24, 10), At(request, "audience", 352, 82, 48, 10)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfLogicalPage page = Assert.Single(result.Document.Pages);

        Assert.Empty(page.Tables);
        Assert.Equal(
            new[] {
                "Alpha project overview continues with details for each audience",
                "Delta project overview continues with details for each audience"
            },
            PdfLogicalReadingOrderAnalysis.Analyze(page)
                .Where(static item => item.Kind == PdfLogicalReadingOrderKind.Paragraph)
                .Select(item => page.Paragraphs[item.SourceIndex].Text));
    }

    [Fact]
    public async Task Document_UsesScaleAwareGapsForLargeOcrTextRuns() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Large", 30, 70, 100, 40),
            At(request, "heading", 160, 70, 130, 40)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfLogicalTextBlock block = Assert.Single(
            result.Document.TextBlocks,
            static block => block.SourceKind == PdfLogicalContentSourceKind.Ocr);

        Assert.Equal("Large heading", block.Text);
        Assert.Empty(result.Document.Tables);
    }

    [Fact]
    public async Task Document_InfersCompactTextOnlyOcrTablesWithoutLanguageTokens() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Pole", 30, 50, 34, 10), At(request, "Stan", 150, 50, 30, 10),
            At(request, "Szukaj", 30, 66, 42, 10), At(request, "Włączone", 150, 66, 54, 10),
            At(request, "Eksport", 30, 82, 46, 10), At(request, "Wyłączone", 150, 82, 60, 10)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfLogicalTable table = Assert.Single(result.Document.Tables);

        Assert.Equal(PdfLogicalContentSourceKind.Ocr, table.SourceKind);
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(new[] { "Eksport", "Wyłączone" }, table.Rows[2]);
    }

    [Fact]
    public async Task Document_InfersOcrTablesIndependentlyOfProviderWordSegmentation() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "一", 30, 50, 7, 10), At(request, "二", 38, 50, 7, 10), At(request, "三", 46, 50, 7, 10), At(request, "四", 54, 50, 7, 10), At(request, "五", 62, 50, 7, 10),
            At(request, "甲", 150, 50, 7, 10), At(request, "乙", 158, 50, 7, 10), At(request, "丙", 166, 50, 7, 10), At(request, "丁", 174, 50, 7, 10), At(request, "戊", 182, 50, 7, 10),
            At(request, "가", 30, 66, 7, 10), At(request, "나", 38, 66, 7, 10), At(request, "다", 46, 66, 7, 10), At(request, "라", 54, 66, 7, 10), At(request, "마", 62, 66, 7, 10),
            At(request, "10", 150, 66, 16, 10),
            At(request, "А", 30, 82, 7, 10), At(request, "Б", 38, 82, 7, 10), At(request, "В", 46, 82, 7, 10), At(request, "Г", 54, 82, 7, 10), At(request, "Д", 62, 82, 7, 10),
            At(request, "20", 150, 82, 16, 10)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        PdfLogicalTable table = Assert.Single(result.Document.Tables);

        Assert.Equal(PdfLogicalContentSourceKind.Ocr, table.SourceKind);
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal("一 二 三 四 五", table.Rows[0][0]);
        Assert.Equal("가 나 다 라 마", table.Rows[1][0]);
        Assert.Equal("А Б В Г Д", table.Rows[2][0]);
        PdfLogicalReadingOrderItem[] readingOrder = PdfLogicalReadingOrderAnalysis
            .Analyze(Assert.Single(result.Document.Pages))
            .ToArray();
        Assert.Single(readingOrder, static item => item.Kind == PdfLogicalReadingOrderKind.Table);
        Assert.DoesNotContain(
            readingOrder,
            static item => item.Kind is PdfLogicalReadingOrderKind.TextBlock or
                PdfLogicalReadingOrderKind.Heading or
                PdfLogicalReadingOrderKind.Paragraph or
                PdfLogicalReadingOrderKind.ListItem);
    }

    [Fact]
    public async Task Document_RetainsOcrTablesWhenCallerReplacesStructuralStages() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "項目", 30, 50, 34, 10), At(request, "値", 150, 50, 30, 10),
            At(request, "甲", 30, 66, 42, 10), At(request, "𝟙𝟚", 150, 66, 54, 10),
            At(request, "乙", 30, 82, 46, 10), At(request, "𝟛𝟜", 150, 82, 60, 10)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider, new PdfOcrMergeOptions {
            ReadOptions = new PdfReadOptions {
                Pipeline = new PdfUnderstandingPipelineOptions {
                    PageSegmentation = new OcrOneRegionSegmentationStage(),
                    ReadingOrder = new OcrIdentityReadingOrderStage()
                }
            }
        });

        PdfLogicalTable table = Assert.Single(result.Document.Tables);
        Assert.Equal(PdfLogicalContentSourceKind.Ocr, table.SourceKind);
        Assert.Equal(3, table.Rows.Count);
    }

    [Fact]
    public void Document_BoundsMaximumSameLineWordProjectionAndHonorsCancellation() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        PdfRecognizedWord[] words = Enumerable.Range(0, 50_000)
            .Select(static index => new PdfRecognizedWord("x", 30, 50, 1, 10, 0.99D, index))
            .ToArray();
        var pageMerge = new PdfOcrPageMergeResult(1, words, 0, 0, Array.Empty<string>(), string.Empty);
        var timer = System.Diagnostics.Stopwatch.StartNew();

        PdfDocumentReadResult enriched = BuildOcrDocument(pdf, new[] { pageMerge }, CancellationToken.None);
        timer.Stop();

        Assert.Single(enriched.TextBlocks, block => block.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.True(
            timer.Elapsed < TimeSpan.FromSeconds(5),
            "Maximum same-line OCR projection exceeded the bounded contract: " + timer.Elapsed + ".");

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() =>
            BuildOcrDocument(pdf, new[] { pageMerge }, cancellation.Token));
    }

    [Fact]
    public void OcrLogicalLines_PreserveProviderLineHierarchyAndReadingSequence() {
        PdfRecognizedWord[] words = {
            new PdfRecognizedWord("الثاني", 20D, 50D, 35D, 10D, 0.9D, 1, "b1", "p1", "l1"),
            new PdfRecognizedWord("الأول", 60D, 50D, 30D, 10D, 0.9D, 0, "b1", "p1", "l1"),
            new PdfRecognizedWord("التالي", 20D, 30D, 35D, 10D, 0.9D, 2, "b1", "p2", "l2")
        };

        IReadOnlyList<PdfOcrLogicalTextLine> lines = PdfOcrLogicalDocumentBuilder.BuildTextLines(words, CancellationToken.None);

        Assert.Equal(2, lines.Count);
        Assert.Equal("الأول الثاني", lines[0].Text);
        Assert.Equal("التالي", lines[1].Text);
    }

    [Fact]
    public void OcrLogicalLines_UseGeometryWhenProviderHierarchyIsAbsent() {
        PdfRecognizedWord[] words = {
            new PdfRecognizedWord("second", 60D, 50D, 35D, 10D, 0.9D, 0),
            new PdfRecognizedWord("first", 20D, 50D, 30D, 10D, 0.9D, 1)
        };

        PdfOcrLogicalTextLine line = Assert.Single(
            PdfOcrLogicalDocumentBuilder.BuildTextLines(words, CancellationToken.None));

        Assert.Equal("first second", line.Text);
    }

    [Fact]
    public void OcrLogicalLines_HonorExplicitDirectionWhenProviderHierarchyIsAbsent() {
        PdfRecognizedWord[] words = {
            new PdfRecognizedWord("left", 20D, 50D, 30D, 10D, 0.9D, 0),
            new PdfRecognizedWord("right", 60D, 50D, 35D, 10D, 0.9D, 1)
        };

        PdfOcrLogicalTextLine line = Assert.Single(PdfOcrLogicalDocumentBuilder.BuildTextLines(
            words,
            PdfReadingDirection.RightToLeft,
            CancellationToken.None));

        Assert.Equal("right left", line.Text);
    }

    [Fact]
    public void OcrLogicalLines_UseGeometryInsteadOfPunctuationToInferWhitespace() {
        PdfRecognizedWord[] words = {
            new PdfRecognizedWord("Alpha", 20D, 50D, 30D, 10D, 0.9D, 0),
            new PdfRecognizedWord(".", 200D, 50D, 4D, 10D, 0.9D, 1),
            new PdfRecognizedWord("Beta", 204D, 50D, 28D, 10D, 0.9D, 2)
        };

        PdfOcrLogicalTextLine line = Assert.Single(
            PdfOcrLogicalDocumentBuilder.BuildTextLines(words, CancellationToken.None));

        Assert.Equal("Alpha .Beta", line.Text);
    }

    [Fact]
    public void OcrLogicalLines_ScopeRepeatedLineIdentifiersByParagraphAndBlock() {
        PdfRecognizedWord[] words = {
            new PdfRecognizedWord("Первый", 20D, 20D, 45D, 10D, 0.9D, 0, "b1", "p1", "1"),
            new PdfRecognizedWord("абзац", 70D, 20D, 35D, 10D, 0.9D, 1, "b1", "p1", "1"),
            new PdfRecognizedWord("Второй", 20D, 50D, 45D, 10D, 0.9D, 2, "b1", "p2", "1"),
            new PdfRecognizedWord("абзац", 70D, 50D, 35D, 10D, 0.9D, 3, "b1", "p2", "1")
        };

        IReadOnlyList<PdfOcrLogicalTextLine> lines = PdfOcrLogicalDocumentBuilder.BuildTextLines(words, CancellationToken.None);

        Assert.Equal(new[] { "Первый абзац", "Второй абзац" }, lines.Select(static line => line.Text));
    }

    [Fact]
    public void OcrLogicalLines_PreserveKnownProviderOrderWhenOneLineLacksHierarchy() {
        PdfRecognizedWord[] words = {
            new PdfRecognizedWord("provider-first", 20D, 80D, 70D, 10D, 0.9D, 0, "b1", "p1", "l1"),
            new PdfRecognizedWord("fallback", 20D, 50D, 50D, 10D, 0.9D, 1),
            new PdfRecognizedWord("provider-second", 20D, 20D, 80D, 10D, 0.9D, 2, "b1", "p2", "l2")
        };

        IReadOnlyList<PdfOcrLogicalTextLine> lines =
            PdfOcrLogicalDocumentBuilder.BuildTextLines(words, CancellationToken.None);

        Assert.Equal(
            new[] { "provider-first", "fallback", "provider-second" },
            lines.Select(static line => line.Text));
    }

    [Fact]
    public void OcrLogicalProjection_TreatsChangedBlockAsParagraphBoundaryEvenWhenParagraphIdRepeats() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        PdfRecognizedWord[] words = {
            new PdfRecognizedWord("Первый", 20D, 20D, 45D, 10D, 0.9D, 0, "b1", "p1", "l1"),
            new PdfRecognizedWord("блок", 70D, 20D, 35D, 10D, 0.9D, 1, "b1", "p1", "l1"),
            new PdfRecognizedWord("Второй", 20D, 34D, 45D, 10D, 0.9D, 2, "b2", "p1", "l1"),
            new PdfRecognizedWord("блок", 70D, 34D, 35D, 10D, 0.9D, 3, "b2", "p1", "l1")
        };
        var pageMerge = new PdfOcrPageMergeResult(1, words, 0, 0, Array.Empty<string>(), string.Empty);

        PdfDocumentReadResult enriched = BuildOcrDocument(pdf, new[] { pageMerge }, CancellationToken.None);

        Assert.Equal(
            new[] { "Первый блок", "Второй блок" },
            enriched.Paragraphs.Select(static paragraph => paragraph.Text));
    }

    [Fact]
    public void OcrLogicalProjection_KeepsHierarchyLinesAtomicAndPlacesFallbackByGeometry() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        PdfRecognizedWord[] words = {
            new PdfRecognizedWord("שלום", 300D, 80D, 40D, 10D, 0.9D, 0, "b1", "p1", "l1"),
            new PdfRecognizedWord("עולם", 20D, 80D, 40D, 10D, 0.9D, 1, "b1", "p1", "l1"),
            new PdfRecognizedWord("fallback", 20D, 10D, 45D, 10D, 0.9D, 2)
        };
        var pageMerge = new PdfOcrPageMergeResult(1, words, 0, 0, Array.Empty<string>(), string.Empty);

        PdfDocumentReadResult enriched = BuildOcrDocument(pdf, new[] { pageMerge }, CancellationToken.None);
        PdfLogicalTextBlock[] blocks = enriched.TextBlocks
            .Where(static block => block.SourceKind == PdfLogicalContentSourceKind.Ocr)
            .ToArray();

        Assert.Equal(2, blocks.Length);
        Assert.Equal("fallback", blocks[0].Text);
        Assert.Equal("שלום עולם", blocks[1].Text);
    }

    [Fact]
    public async Task OcrMerge_DiscardsOversizedHierarchyIdentifiersWithoutAbortingRecognition() {
        string oversized = new string(' ', 256) + "x";
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "text", 20D, 20D, 10D, 10D, blockId: oversized, paragraphId: oversized, lineId: oversized)
        }));

        PdfOcrPageMergeResult page = Assert.Single((await PdfDocument.Load(pdf).ReadWithOcrAsync(provider)).Pages);
        PdfRecognizedWord word = Assert.Single(page.Words);

        Assert.Null(word.BlockId);
        Assert.Null(word.ParagraphId);
        Assert.Null(word.LineId);
        Assert.Contains(page.Diagnostics, diagnostic => diagnostic.StartsWith("ocr-hierarchy-id-limit:", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Document_ProjectsOcrHeadingsListsAndParagraphsThroughSemanticAdapters() {
        byte[] pdf = PdfDocument.Create().Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120).ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Quarterly", 30, 90, 86, 24), At(request, "review", 124, 90, 54, 24),
            At(request, "Summary", 30, 140, 54, 12), At(request, "narrative", 90, 140, 60, 12),
            At(request, "-", 30, 170, 6, 12), At(request, "Ready", 42, 170, 38, 12),
            At(request, "Closing", 30, 200, 42, 12), At(request, "note", 78, 200, 28, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);
        Assert.Contains(result.Document.Headings, heading => heading.Text == "Quarterly review");
        Assert.Contains(result.Document.ListItems, item => item.Marker == "-" && item.Text == "Ready");
        Assert.Contains(result.Document.Paragraphs, paragraph => paragraph.Text.Contains("Summary narrative", StringComparison.Ordinal));

        using (OfficeIMO.Word.WordDocument word = result.Document.ToWordDocument()) {
            Assert.True(word.ToBytes().Length > 100);
        }
        string html = result.Document.ToHtml();
        Assert.Contains("Quarterly review", html, StringComparison.Ordinal);
        Assert.Contains("Ready", html, StringComparison.Ordinal);
        Assert.Contains("Summary narrative", result.Document.ToRtfDocument().ToRtf(), StringComparison.Ordinal);
        PdfOdtConversionResult odtResult = result.Document.ToOdtDocumentResult();
        Assert.Contains(odtResult.Value.Paragraphs, paragraph => paragraph.Text.Contains("Summary narrative", StringComparison.Ordinal));
    }

    [Fact]
    public async Task Document_DoesNotRemoveOcrDecimalParagraphsAsListItems() {
        byte[] pdf = PdfDocument.Create().Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120).ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "1037.25", 30, 140, 48, 12), At(request, "total", 84, 140, 30, 12),
            At(request, "1.2.", 30, 170, 24, 12), At(request, "Nested", 60, 170, 42, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);

        Assert.DoesNotContain(result.Document.ListItems,
            item => item.Text.Contains("1037.25", StringComparison.Ordinal));
        Assert.Contains(result.Document.Paragraphs,
            paragraph => paragraph.Text.Contains("1037.25 total", StringComparison.Ordinal));
        Assert.Contains(result.Document.ListItems,
            item => item.Marker == "1.2" && item.Text == "Nested");
    }

    [Fact]
    public async Task Document_AlwaysProjectsAcceptedOcrThroughCanonicalAnalysis() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Native anchor")).ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] { At(request, "OCR", 30, 250, 30, 12) }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(provider);

        Assert.Single(result.Pages[0].Words);
        PdfLogicalTextBlock block = Assert.Single(result.Document.TextBlocks, block => block.SourceKind == PdfLogicalContentSourceKind.Ocr);
        PdfUnderstandingSemanticElement element = Assert.Single(
            result.Document.Pages[0].Analysis.Elements,
            element => element.Region.Lines.Any(line => line.SourceKind == PdfLogicalContentSourceKind.Ocr));
        Assert.Contains("OCR", block.Text, StringComparison.Ordinal);
        Assert.Contains("OCR", element.Region.Text, StringComparison.Ordinal);
        Assert.True(result.HasAcceptedOcrContent);
    }

    [Fact]
    public async Task MakeSearchableAsync_WritesGeometryAlignedInvisibleTextWithoutVisualChanges() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Searchable", 42, 160, 88, 14),
            At(request, "document", 138, 160, 70, 14),
            At(request, "Zażółć", 42, 200, 58, 14)
        }, diagnostics: null, provider: "fixture", model: "fixture-v1", language: "eng"));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).MakeSearchableAsync(provider);
        byte[] searchable = result.Document.ToBytes();

        Assert.True(result.WasModified);
        Assert.Equal(3, result.AddedWordCount);
        Assert.Equal(new[] { 1 }, result.ModifiedPages);
        Assert.Contains("Searchable document", PdfReadDocument.Open(searchable).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Zażółć", PdfReadDocument.Open(searchable).ExtractText(), StringComparison.Ordinal);
        PdfPageInteractionMap interactions = PdfPageInteractionMap.Create(
            searchable,
            1,
            new PdfPageInteractionOptions { IncludeInvisibleText = true });
        PdfPageInteractionRegion[] word = interactions.TextRegions.Take("Searchable".Length).ToArray();
        Assert.Equal("Searchable", string.Concat(word.Select(static region => region.Text)));
        Assert.Equal(42D, word[0].Quad.Left, 2);
        Assert.Equal(88D, word[word.Length - 1].Quad.Right - word[0].Quad.Left, 2);
        Assert.Empty(PdfPageInteractionMap.Create(searchable, 1).TextRegions);
        Assert.True(PdfVisualComparer.Compare(source, searchable).IsMatch);
        Assert.Equal("fixture", result.Ocr.Pages[0].Provider);
    }

    [Fact]
    public async Task MakeSearchableAsync_PreservesProviderLogicalOrderForRightToLeftWords() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "שלום", 150, 160, 48, 14, blockId: "b1", paragraphId: "p1", lineId: "l1"),
            At(request, "עולם", 100, 160, 42, 14, blockId: "b1", paragraphId: "p1", lineId: "l1")
        }, diagnostics: null, provider: "fixture", model: "fixture-v1", language: "heb"));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).MakeSearchableAsync(provider);
        byte[] searchable = result.Document.ToBytes();

        Assert.Contains("שלום עולם", result.Ocr.Pages[0].Text, StringComparison.Ordinal);
        Assert.Contains(result.Ocr.Document.TextBlocks,
            block => block.SourceKind == PdfLogicalContentSourceKind.Ocr && block.Text == "שלום עולם");
        Assert.DoesNotContain(
            result.Ocr.Document.Pages[0].Analysis.ReadingOrderEvidence.SelectMany(static item => item.Evidence),
            static evidence => evidence.Code == "reading-order.geometry-consistent" ||
                evidence.Code == "reading-order.geometry-conflict");
        string decodedContent = string.Join(
            Environment.NewLine,
            PdfDocument.Load(searchable).Debug(new PdfDebuggerOptions {
                IncludeDecodedStreamPreviews = true,
                MaxDecodedStreamPreviewBytes = 64 * 1024
            }).Objects.Select(static item => item.DecodedStreamPreview ?? string.Empty));
        int firstWord = decodedContent.IndexOf(PdfSyntaxEscaper.TextString("שלום"), StringComparison.Ordinal);
        int secondWord = decodedContent.IndexOf(PdfSyntaxEscaper.TextString("עולם"), StringComparison.Ordinal);
        Assert.True(firstWord >= 0, "The searchable layer did not contain the first logical word.");
        Assert.True(secondWord > firstWord, "The searchable layer reversed the provider's right-to-left logical word sequence.");
    }

    [Fact]
    public async Task MakeSearchableAsync_PreservesProviderSequenceAcrossLinesAndColumns() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "LeftTop", 30, 80, 54, 12, blockId: "b1", paragraphId: "p1", lineId: "l1"),
            At(request, "LeftBottom", 30, 120, 66, 12, blockId: "b1", paragraphId: "p1", lineId: "l2"),
            At(request, "RightTop", 300, 80, 60, 12, blockId: "b1", paragraphId: "p1", lineId: "l3"),
            At(request, "RightBottom", 300, 120, 72, 12, blockId: "b1", paragraphId: "p1", lineId: "l4")
        }));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).MakeSearchableAsync(provider);
        string mergedText = result.Ocr.Pages[0].Text;
        Assert.Equal(
            PdfDocumentReadResult.GetCanonicalPageText(result.Ocr.Document.Pages[0]),
            mergedText);
        string[] expected = { "LeftTop", "LeftBottom", "RightTop", "RightBottom" };
        int previousText = -1;
        for (int index = 0; index < expected.Length; index++) {
            int currentText = mergedText.IndexOf(expected[index], StringComparison.Ordinal);
            Assert.True(currentText > previousText, $"The OCR page text did not preserve provider sequence at '{expected[index]}'.");
            previousText = currentText;
        }
        string decodedContent = string.Join(
            Environment.NewLine,
            result.Document.Debug(new PdfDebuggerOptions {
                IncludeDecodedStreamPreviews = true,
                MaxDecodedStreamPreviewBytes = 64 * 1024
            }).Objects.Select(static item => item.DecodedStreamPreview ?? string.Empty));

        int previous = -1;
        for (int index = 0; index < expected.Length; index++) {
            int current = decodedContent.IndexOf(PdfSyntaxEscaper.TextString(expected[index]), StringComparison.Ordinal);
            Assert.True(current > previous, $"The searchable layer did not preserve provider sequence at '{expected[index]}'.");
            previous = current;
        }
    }

    [Fact]
    public async Task MakeSearchableAsync_DeduplicatesSelectedPhysicalPagesBeforeRecognition() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Once", 42, 160, 40, 14)
        }));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).MakeSearchableAsync(provider, new PdfOcrMergeOptions {
            ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(1, 1) }
        });

        Assert.Equal(1, provider.CallCount);
        Assert.Equal(1, result.AddedWordCount);
        Assert.Equal(new[] { 1 }, result.ModifiedPages);
        Assert.Contains("Once", PdfReadDocument.Open(result.Document.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task MakeSearchableAsync_DoesNotDuplicateAnExistingInvisibleSearchLayer() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Already searchable", 42, 160, 120, 14)
        }));

        PdfSearchableOcrResult first = await PdfDocument.Load(source).MakeSearchableAsync(provider);
        PdfSearchableOcrResult second = await first.Document.MakeSearchableAsync(provider);

        Assert.True(first.WasModified);
        Assert.False(second.WasModified);
        Assert.Equal(0, second.AddedWordCount);
        Assert.Equal(1, second.Ocr.Pages[0].RejectedNativeOverlapCount);
        Assert.Empty(second.ModifiedPages);
        Assert.Same(first.Document, second.Document);
    }

    [Fact]
    public async Task MakeSearchableAsync_LeavesDocumentUnchangedWhenNoWordsAreAccepted() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Already searchable")).ToBytes();
        var provider = new StubOcrEngine(_ => Result(Array.Empty<OcrTextSpan>()));
        PdfDocument document = PdfDocument.Load(source);

        PdfSearchableOcrResult result = await document.MakeSearchableAsync(provider);

        Assert.False(result.WasModified);
        Assert.Equal(0, result.AddedWordCount);
        Assert.Empty(result.ModifiedPages);
        Assert.Same(document, result.Document);
        Assert.Equal(source, result.Document.ToBytes());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public async Task MakeSearchableAsync_PreservesVisualCoordinatesAcrossCropAndRotation(int rotation) {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        source = PdfPageEditor.SetCropBox(source, 20, 40, 500, 700);
        source = PdfPageEditor.RotatePages(source, rotation);
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "Rotated", 24, 180, 64, 12)
        }));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).MakeSearchableAsync(provider);
        PdfPageInteractionMap interactions = PdfPageInteractionMap.Create(
            result.Document.ToBytes(),
            1,
            new PdfPageInteractionOptions { IncludeInvisibleText = true });
        PdfPageInteractionRegion[] word = interactions.TextRegions.ToArray();

        Assert.Equal("Rotated", string.Concat(word.Select(static region => region.Text)));
        Assert.Equal(24D, word[0].Quad.Left, 2);
        Assert.Equal(180D, word[0].Quad.Top, 2);
        Assert.Equal(64D, word[word.Length - 1].Quad.Right - word[0].Quad.Left, 2);
        Assert.True(PdfVisualComparer.Compare(source, result.Document.ToBytes()).IsMatch);
    }

    private static OcrTextSpan At(
        OcrRequest request,
        string text,
        double x,
        double y,
        double width,
        double height,
        double confidence = 0.95D,
        string? blockId = null,
        string? paragraphId = null,
        string? lineId = null) =>
        new OcrTextSpan {
            Level = OcrTextSpanLevel.Word,
            Text = text,
            Confidence = confidence,
            CoordinateUnit = OcrCoordinateUnit.Points,
            Region = new OcrRegion { X = x, Y = y, Width = width, Height = height },
            BlockId = blockId,
            ParagraphId = paragraphId,
            LineId = lineId
        };

    private static double Scale(OcrRequest request) {
        if (request.PixelWidth is not > 0 || request.Region == null || request.Region.Width <= 0D) {
            throw new InvalidOperationException("The PDF OCR request did not expose usable pixel and point dimensions.");
        }
        return request.PixelWidth.Value / request.Region.Width;
    }

    [Fact]
    public async Task OcrMerge_AcceptsExactlyFilledRawHierarchyCharacterBudget() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "text", 20D, 20D, 10D, 10D, blockId: "b", paragraphId: "p", lineId: "l")
        }));

        PdfOcrPageMergeResult page = Assert.Single((await PdfDocument.Load(pdf).ReadWithOcrAsync(
            provider,
            new PdfOcrMergeOptions { MaxOcrHierarchyCharactersPerPage = 3 })).Pages);

        PdfRecognizedWord word = Assert.Single(page.Words);
        Assert.Equal("b", word.BlockId);
        Assert.Equal("p", word.ParagraphId);
        Assert.Equal("l", word.LineId);
    }

    [Fact]
    public async Task OcrMerge_RejectsRawHierarchyCharactersOnePastAggregateBudget() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrEngine(request => Result(new[] {
            At(request, "text", 20D, 20D, 10D, 10D, blockId: "b", paragraphId: "p", lineId: "l")
        }));

        PdfReadLimitException exception = await Assert.ThrowsAsync<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf).ReadWithOcrAsync(
                provider,
                new PdfOcrMergeOptions { MaxOcrHierarchyCharactersPerPage = 2 }));

        Assert.Equal(PdfReadLimitKind.OcrArtifacts, exception.Kind);
        Assert.Equal(2, exception.Limit);
        Assert.Equal(3, exception.Actual);
    }

    private static OcrTextSpan Word(
        string text,
        double x,
        double y,
        double width,
        double height,
        double confidence,
        string? blockId = null,
        string? paragraphId = null,
        string? lineId = null) =>
        new OcrTextSpan {
            Level = OcrTextSpanLevel.Word,
            Text = text,
            Confidence = confidence,
            CoordinateUnit = OcrCoordinateUnit.Pixels,
            Region = new OcrRegion { X = x, Y = y, Width = width, Height = height },
            BlockId = blockId,
            ParagraphId = paragraphId,
            LineId = lineId
        };

    private static OcrResult Result(
        IEnumerable<OcrTextSpan> spans,
        IEnumerable<string>? diagnostics = null,
        string? provider = null,
        string? model = null,
        string? language = null) =>
        new OcrResult {
            Spans = spans.ToArray(),
            Diagnostics = (diagnostics ?? Array.Empty<string>())
                .Select(static message => new OcrDiagnostic { Message = message })
                .ToArray(),
            Provider = provider,
            Model = model,
            Language = language
        };

    private static PdfDocumentReadResult BuildOcrDocument(
        byte[] pdf,
        IReadOnlyList<PdfOcrPageMergeResult> pages,
        CancellationToken cancellationToken) {
        PdfReadDocument source = PdfReadDocument.Open(pdf, null, cancellationToken);
        var layoutOptions = new PdfTextLayoutOptions();
        var pipelineOptions = new PdfUnderstandingPipelineOptions();
        PdfDocumentReadResult native = PdfDocumentReadEngine.Read(
            source,
            new PdfReadOptions {
                Profile = PdfReadProfile.Structured,
                LayoutOptions = layoutOptions,
                Pipeline = pipelineOptions
            },
            out IReadOnlyList<PdfUnderstandingPageResult> nativePageAnalyses,
            cancellationToken);
        return PdfOcrLogicalDocumentBuilder.Build(
            source,
            native,
            nativePageAnalyses,
            pages,
            layoutOptions,
            pipelineOptions,
            cancellationToken);
    }

    private sealed class ChangingIdentityOcrEngine : IOcrEngine {
        private int _idReadCount;
        private int _capabilitiesReadCount;
        private int _callCount;

        internal int IdReadCount => Volatile.Read(ref _idReadCount);
        internal int CapabilitiesReadCount => Volatile.Read(ref _capabilitiesReadCount);
        internal int CallCount => Volatile.Read(ref _callCount);

        public string Id => Interlocked.Increment(ref _idReadCount) == 1
            ? "snapshot-engine"
            : new string('x', OcrEngineRunner.MaximumEngineIdCharacters + 1);

        public OcrEngineCapabilities Capabilities {
            get {
                Interlocked.Increment(ref _capabilitiesReadCount);
                return new OcrEngineCapabilities {
                    SupportedMediaTypes = new[] { "image/png" },
                    SupportsConcurrentRequests = true
                };
            }
        }

        public Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
            Interlocked.Increment(ref _callCount);
            return Task.FromResult(new OcrResult());
        }
    }

    private sealed class StubOcrEngine : IOcrEngine {
        private readonly Func<OcrRequest, OcrResult> _response;
        private readonly string _id;
        public StubOcrEngine(Func<OcrRequest, OcrResult> response, string id = "fixture") {
            _response = response;
            _id = id;
        }
        public int CallCount { get; private set; }
        public OcrRequest? LastRequest { get; private set; }
        public string Id => _id;
        public OcrEngineCapabilities Capabilities { get; } = new OcrEngineCapabilities {
            SupportedMediaTypes = new[] { "image/png" },
            SupportsWordSpans = true,
            SupportsConfidence = true
        };
        public Task<OcrResult> RecognizeAsync(OcrRequest request, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            CallCount++;
            LastRequest = request;
            return Task.FromResult(_response(request));
        }
    }

    private sealed class OcrHeadingClassificationStage : IPdfSemanticClassificationStage {
        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingRegion> orderedRegions) =>
            orderedRegions
                .Select(static region => new PdfUnderstandingSemanticElement(
                    region,
                    PdfUnderstandingSemanticKind.Heading,
                    level: 1))
                .ToArray();
    }

    private sealed class OcrOneRegionSegmentationStage : IPdfPageSegmentationStage {
        public IReadOnlyList<PdfUnderstandingRegion> Segment(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingLine> lines) =>
            lines.Count == 0
                ? Array.Empty<PdfUnderstandingRegion>()
                : new[] { new PdfUnderstandingRegion(lines.Reverse().ToArray()) };
    }

    private sealed class OcrIdentityReadingOrderStage : IPdfReadingOrderStage {
        public IReadOnlyList<PdfUnderstandingRegion> Order(
            PdfUnderstandingPageContext context,
            IReadOnlyList<PdfUnderstandingRegion> regions) => regions;
    }
}
