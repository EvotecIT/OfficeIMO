using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            new PdfOcrWord("Native", 150, 140, 100, 30, 0.99),
            new PdfOcrWord("Scanned", 150, 400, 120, 32, 0.95),
            new PdfOcrWord("Weak", 300, 400, 80, 30, 0.2),
            new PdfOcrWord("Outside", request.PixelWidth, 0, 20, 20, 0.99)
        }, new[] { "provider-proof" }, provider: "fixture", model: "fixture-v1", language: "eng"));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);
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
        Assert.Equal("fixture", page.Provider);
        Assert.Equal("fixture-v1", page.Model);
        Assert.Equal("eng", page.Language);
        Assert.Contains(page.Diagnostics, diagnostic => diagnostic.StartsWith("InvalidWordGeometry:", StringComparison.Ordinal));
        Assert.Contains("Native text", page.Text, StringComparison.Ordinal);
        Assert.Contains("Scanned", page.Text, StringComparison.Ordinal);
        Assert.Same(result.NativeDocument.Pages[0], Assert.Single(result.NativeDocument.Pages));
        PdfLogicalTextBlock ocrBlock = Assert.Single(result.EnrichedDocument.TextBlocks, block => block.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.Equal("Scanned", ocrBlock.Text);
        Assert.NotNull(ocrBlock.VisualBounds);
        Assert.InRange(ocrBlock.Confidence, 0.94D, 0.96D);
        Assert.True(result.HasAcceptedOcrContent);
        Assert.Equal(1, result.AcceptedWordCount);
    }

    [Fact]
    public async Task EnrichedDocument_MergesNativeAndOcrBlocksInVisualReadingOrder() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Native follows OCR"))
            .ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "OCR first", 30, 5, 70, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);
        PdfLogicalPage page = Assert.Single(result.EnrichedDocument.Pages);

        Assert.Equal("OCR first", page.TextBlocks[0].Text);
        Assert.Equal(
            page.TextBlocks.Select(static block => block.Text),
            page.Elements.OfType<PdfLogicalTextBlock>().Select(static block => block.Text));
        Assert.True(page.TextBlocks.ToList().FindIndex(static block =>
            block.Text.Contains("Native follows OCR", StringComparison.Ordinal)) > 0);
        Assert.True(result.EnrichedDocument.Text.IndexOf("OCR first", StringComparison.Ordinal) <
            result.EnrichedDocument.Text.IndexOf("Native follows OCR", StringComparison.Ordinal));
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
    public async Task RecognizeAndMergeAsync_UsesCanonicalStructuredReadForNativeEvidence() {
        byte[] pdf = PdfDocument.Create()
            .TaggedPdfCatalogMarkers()
            .H1("Native structured heading")
            .ToBytes();
        var provider = new StubOcrProvider(_ => new PdfOcrResponse(Array.Empty<PdfOcrWord>()));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);
        PdfUnderstandingPageResult analysis = Assert.Single(result.NativeDocument.Pages).Analysis;

        Assert.Equal(PdfReadProfile.Structured, result.NativeDocument.Profile);
        Assert.Contains(analysis.Elements, element =>
            element.Kind == PdfUnderstandingSemanticKind.Heading &&
            element.Region.Text.Contains("Native structured heading", StringComparison.Ordinal) &&
            element.Evidence.Any(static evidence => evidence.Code == "semantic.tagged-pdf-role"));
    }

    [Fact]
    public void OcrSemanticPipeline_InheritsThePublicOcrPageLimit() {
        var options = new PdfOcrMergeOptions { MaxPages = 1_500 };

        PdfUnderstandingPipelineOptions pipeline = PdfOcr.CreateUnderstandingPipelineOptions(options);

        Assert.Equal(options.MaxPages, pipeline.MaxPages);
    }

    [Fact]
    public async Task RecognizeAndMergeAsync_PreservesDuplicateCallerOrderedPages() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "Repeated", 30, 90, 50, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider, new PdfOcrMergeOptions {
            Selection = PdfPageSelection.From(1, 1),
            DetectAlignedTables = false
        });

        Assert.Equal(2, provider.CallCount);
        Assert.Equal(new[] { 1, 1 }, result.Pages.Select(static page => page.PageNumber));
        Assert.Equal(new[] { 1, 1 }, result.EnrichedDocument.Pages.Select(static page => page.PageNumber));
        Assert.All(result.EnrichedDocument.Pages, page =>
            Assert.Single(page.TextBlocks, block => block.SourceKind == PdfLogicalContentSourceKind.Ocr));
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

    [Fact]
    public async Task RecognizeAndMergeAsync_BoundsNativeOverlapWork() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First native block"))
            .Paragraph(paragraph => paragraph.Text("Second native block"))
            .ToBytes();
        var provider = new StubOcrProvider(_ => new PdfOcrResponse(new[] {
            new PdfOcrWord("scanned", 10, 10, 10, 10, 0.9)
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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            new PdfOcrWord(
                "Footer",
                left * request.Scale,
                top * request.Scale,
                (right - left) * request.Scale,
                (bottom - top) * request.Scale,
                0.99D)
        }));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).Ocr.MakeSearchableAsync(provider);

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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            new PdfOcrWord("Native", left * request.Scale, top * request.Scale,
                (right - left) * request.Scale, (bottom - top) * request.Scale, 0.99D)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(source).Ocr.ReadAsync(provider);

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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            new PdfOcrWord("Scanned", glyph.Left * request.Scale, glyph.Top * request.Scale,
                glyph.Width * 3D * request.Scale, glyph.Height * request.Scale, 0.99D)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(source).Ocr.ReadAsync(provider);

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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            new PdfOcrWord("Scanned", left * request.Scale, top * request.Scale,
                (right - left) * request.Scale, (bottom - top) * request.Scale, 0.99D)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(source).Ocr.ReadAsync(provider);

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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            new PdfOcrWord("Scanned", ocrLeft * request.Scale, firstLogicalGlyph.Top * request.Scale,
                24D * request.Scale, firstLogicalGlyph.Height * request.Scale, 0.99D)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(source).Ocr.ReadAsync(provider);

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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            new PdfOcrWord(
                "duplicate",
                nativeGlyph.Left * request.Scale,
                nativeGlyph.Top * request.Scale,
                nativeGlyph.Width * request.Scale,
                nativeGlyph.Height * request.Scale,
                0.99)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);
        PdfOcrPageMergeResult page = Assert.Single(result.Pages);

        Assert.Equal(map.Width, provider.LastRequest!.PageWidth, 3);
        Assert.Equal(map.Height, provider.LastRequest.PageHeight, 3);
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
    public async Task EnrichedDocument_PreservesOcrVisualGeometryAcrossPageRotation(int rotation) {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Native anchor")).ToBytes();
        pdf = PdfPageEditor.SetCropBox(pdf, 20, 40, 500, 700);
        pdf = PdfPageEditor.RotatePages(pdf, rotation);
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "OCR", 24, 180, 42, 12),
            At(request, "geometry", 72, 180, 64, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);
        PdfLogicalTextBlock block = Assert.Single(result.EnrichedDocument.TextBlocks, candidate => candidate.SourceKind == PdfLogicalContentSourceKind.Ocr);
        PdfLogicalVisualBounds bounds = Assert.IsType<PdfLogicalVisualBounds>(block.VisualBounds);
        Assert.Equal(24D, bounds.Left, 3);
        Assert.Equal(180D, bounds.Top, 3);
        Assert.Equal("OCR geometry", block.Text);
        PdfLogicalPage logicalPage = result.EnrichedDocument.Pages[0];
        PdfLogicalReadingOrderItem entry = Assert.Single(
            PdfLogicalReadingOrderAnalysis.Analyze(logicalPage),
            candidate => candidate.Kind == PdfLogicalReadingOrderKind.Paragraph &&
                logicalPage.Paragraphs[candidate.SourceIndex].Text.Contains("OCR geometry", StringComparison.Ordinal));
        Assert.Equal(24D, entry.Left, 3);
        Assert.Equal(180D, entry.Top, 3);
    }

    [Fact]
    public async Task EnrichedDocument_InfersOnlyRepeatedAlignedOcrRowsAsEditableTable() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Native anchor")).ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "Item", 30, 220, 34, 12), At(request, "Value", 150, 220, 38, 12),
            At(request, "Alpha", 30, 245, 38, 12), At(request, "10", 150, 245, 18, 12),
            At(request, "Beta", 30, 270, 32, 12), At(request, "20", 150, 270, 18, 12),
            At(request, "Narrative", 30, 320, 58, 12), At(request, "continues", 94, 320, 52, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);
        PdfLogicalTable table = Assert.Single(result.EnrichedDocument.Tables, candidate => candidate.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.Equal("OcrAlignedColumns", table.DetectionKind);
        Assert.Equal(PdfTableCoordinateSpace.VisualTopLeft, table.CoordinateSpace);
        PdfUnderstandingTableCandidate candidate = Assert.Single(
            result.EnrichedDocument.Pages[0].Analysis.TableCandidates,
            candidate => candidate.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.Equal(table.Rows.Select(static row => row.ToArray()), candidate.Rows.Select(static row => row.ToArray()));
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(new[] { "Alpha", "10" }, table.Rows[1]);
        Assert.DoesNotContain(table.Rows, row => row.Contains("Narrative"));

        PdfExcelTableImportResult excelResult = result.EnrichedDocument.ImportTablesToExcelDocumentResult();
        using (excelResult.Value) Assert.Single(excelResult.Report.Entries);
        PdfOdsConversionResult odsResult = result.EnrichedDocument.ToOdsDocumentResult();
        Assert.Single(odsResult.Value.Sheets);

        PdfPowerPointConversionResult presentationResult = result.EnrichedDocument
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateEditableContent());
        using (presentationResult.Value) {
            Assert.Single(presentationResult.Report.EditablePages);
            Assert.Equal(1, presentationResult.Report.EditablePages[0].TableCount);
        }

        string html = result.EnrichedDocument.ToHtml();
        Assert.Equal(1, html.Split(new[] { "Alpha" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void EnrichedDocument_DoesNotDuplicateNativeTableWithOverlappingOcrCandidate() {
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

        PdfLogicalPage enriched = page.WithOcrContent(
            Array.Empty<PdfLogicalTextBlock>(),
            Array.Empty<PdfLogicalHeading>(),
            Array.Empty<PdfLogicalParagraph>(),
            Array.Empty<PdfLogicalListItem>(),
            new[] { ocr });

        Assert.Single(enriched.Tables);
        Assert.Single(enriched.Analysis.TableCandidates);
        Assert.All(enriched.Tables, table => Assert.Equal(PdfLogicalContentSourceKind.Native, table.SourceKind));
    }

    [Fact]
    public void EnrichedDocument_PreservesRicherOverlappingOcrTable() {
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

        PdfLogicalPage enriched = page.WithOcrContent(
            Array.Empty<PdfLogicalTextBlock>(),
            Array.Empty<PdfLogicalHeading>(),
            Array.Empty<PdfLogicalParagraph>(),
            Array.Empty<PdfLogicalListItem>(),
            new[] { ocr });

        PdfLogicalTable ocrTable = Assert.Single(enriched.Tables, table => table.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.Equal(4, ocrTable.Rows.Count);
        Assert.Contains(ocrTable.Rows, row => row.SequenceEqual(new[] { "Gamma", "30" }));
        Assert.Equal(2, enriched.Analysis.TableCandidates.Count);
    }

    [Fact]
    public async Task EnrichedDocument_KeepsAlignedNarrativeColumnsOutOfTableInference() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "Alpha", 30, 50, 30, 10), At(request, "project", 65, 50, 42, 10), At(request, "overview", 112, 50, 48, 10),
            At(request, "Delta", 300, 50, 30, 10), At(request, "project", 335, 50, 42, 10), At(request, "overview", 382, 50, 48, 10),
            At(request, "continues", 30, 66, 48, 10), At(request, "with", 83, 66, 24, 10), At(request, "details", 112, 66, 36, 10),
            At(request, "continues", 300, 66, 48, 10), At(request, "with", 353, 66, 24, 10), At(request, "details", 382, 66, 36, 10),
            At(request, "for", 30, 82, 18, 10), At(request, "each", 53, 82, 24, 10), At(request, "audience", 82, 82, 48, 10),
            At(request, "for", 300, 82, 18, 10), At(request, "each", 323, 82, 24, 10), At(request, "audience", 352, 82, 48, 10)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);
        PdfLogicalPage page = Assert.Single(result.EnrichedDocument.Pages);

        Assert.Empty(page.Tables);
        Assert.Equal(6, page.Paragraphs.Count);
        Assert.Equal(
            new[] {
                "Alpha project overview",
                "continues with details",
                "for each audience",
                "Delta project overview",
                "continues with details",
                "for each audience"
            },
            PdfLogicalReadingOrderAnalysis.Analyze(page)
                .Where(static item => item.Kind == PdfLogicalReadingOrderKind.Paragraph)
                .Select(item => page.Paragraphs[item.SourceIndex].Text));
    }

    [Fact]
    public async Task EnrichedDocument_InfersCompactTextOnlyOcrTablesWithoutLanguageTokens() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "Pole", 30, 50, 34, 10), At(request, "Stan", 150, 50, 30, 10),
            At(request, "Szukaj", 30, 66, 42, 10), At(request, "Włączone", 150, 66, 54, 10),
            At(request, "Eksport", 30, 82, 46, 10), At(request, "Wyłączone", 150, 82, 60, 10)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);
        PdfLogicalTable table = Assert.Single(result.EnrichedDocument.Tables);

        Assert.Equal(PdfLogicalContentSourceKind.Ocr, table.SourceKind);
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(new[] { "Eksport", "Wyłączone" }, table.Rows[2]);
    }

    [Fact]
    public void EnrichedDocument_BoundsMaximumSameLineWordProjectionAndHonorsCancellation() {
        byte[] pdf = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        PdfDocumentReadResult native = PdfDocumentReadResult.Load(pdf);
        PdfRecognizedWord[] words = Enumerable.Range(0, 50_000)
            .Select(static index => new PdfRecognizedWord("x", 30, 50, 1, 10, 0.99D, index))
            .ToArray();
        var pageMerge = new PdfOcrPageMergeResult(1, words, 0, 0, Array.Empty<string>(), string.Empty);
        var timer = System.Diagnostics.Stopwatch.StartNew();

        PdfDocumentReadResult enriched = PdfOcrLogicalDocumentBuilder.Build(
            native,
            new[] { pageMerge },
            new PdfOcrMergeOptions(),
            CancellationToken.None);
        timer.Stop();

        Assert.Single(enriched.TextBlocks, block => block.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.True(
            timer.Elapsed < TimeSpan.FromSeconds(5),
            "Maximum same-line OCR projection exceeded the bounded contract: " + timer.Elapsed + ".");

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => PdfOcrLogicalDocumentBuilder.Build(
            native,
            new[] { pageMerge },
            new PdfOcrMergeOptions(),
            cancellation.Token));
    }

    [Fact]
    public async Task EnrichedDocument_ProjectsOcrHeadingsListsAndParagraphsThroughSemanticAdapters() {
        byte[] pdf = PdfDocument.Create().Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120).ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "Quarterly", 30, 90, 86, 24), At(request, "review", 124, 90, 54, 24),
            At(request, "Summary", 30, 140, 54, 12), At(request, "narrative", 90, 140, 60, 12),
            At(request, "-", 30, 170, 6, 12), At(request, "Ready", 42, 170, 38, 12),
            At(request, "Closing", 30, 200, 42, 12), At(request, "note", 78, 200, 28, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);
        Assert.Contains(result.EnrichedDocument.Headings, heading => heading.Text == "Quarterly review");
        Assert.Contains(result.EnrichedDocument.ListItems, item => item.Marker == "-" && item.Text == "Ready");
        Assert.Contains(result.EnrichedDocument.Paragraphs, paragraph => paragraph.Text.Contains("Summary narrative", StringComparison.Ordinal));

        using (OfficeIMO.Word.WordDocument word = result.EnrichedDocument.ToWordDocument()) {
            Assert.True(word.ToBytes().Length > 100);
        }
        string html = result.EnrichedDocument.ToHtml();
        Assert.Contains("Quarterly review", html, StringComparison.Ordinal);
        Assert.Contains("Ready", html, StringComparison.Ordinal);
        Assert.Contains("Summary narrative", result.EnrichedDocument.ToRtfDocument().ToRtf(), StringComparison.Ordinal);
        PdfOdtConversionResult odtResult = result.EnrichedDocument.ToOdtDocumentResult();
        Assert.Contains(odtResult.Value.Paragraphs, paragraph => paragraph.Text.Contains("Summary narrative", StringComparison.Ordinal));
    }

    [Fact]
    public async Task EnrichedDocument_DoesNotRemoveOcrDecimalParagraphsAsListItems() {
        byte[] pdf = PdfDocument.Create().Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120).ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "1037.25", 30, 140, 48, 12), At(request, "total", 84, 140, 30, 12),
            At(request, "1.2.", 30, 170, 24, 12), At(request, "Nested", 60, 170, 42, 12)
        }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider);

        Assert.DoesNotContain(result.EnrichedDocument.ListItems,
            item => item.Text.Contains("1037.25", StringComparison.Ordinal));
        Assert.Contains(result.EnrichedDocument.Paragraphs,
            paragraph => paragraph.Text.Contains("1037.25 total", StringComparison.Ordinal));
        Assert.Contains(result.EnrichedDocument.ListItems,
            item => item.Marker == "1.2" && item.Text == "Nested");
    }

    [Fact]
    public async Task EnrichedDocument_CanDisableSemanticProjectionWithoutChangingMergeEvidence() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Native anchor")).ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] { At(request, "OCR", 30, 250, 30, 12) }));

        PdfOcrMergeResult result = await PdfDocument.Load(pdf).Ocr.ReadAsync(provider, new PdfOcrMergeOptions {
            BuildEnrichedLogicalDocument = false
        });

        Assert.Single(result.Pages[0].Words);
        Assert.Same(result.NativeDocument, result.EnrichedDocument);
        Assert.DoesNotContain(result.EnrichedDocument.TextBlocks, block => block.SourceKind == PdfLogicalContentSourceKind.Ocr);
        Assert.True(result.HasAcceptedOcrContent);
    }

    [Fact]
    public async Task MakeSearchableAsync_WritesGeometryAlignedInvisibleTextWithoutVisualChanges() {
        byte[] source = PdfDocument.Create()
            .Image(PdfPngTestImages.CreateRgbPng(245, 245, 245), 220, 120)
            .ToBytes();
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "Searchable", 42, 160, 88, 14),
            At(request, "document", 138, 160, 70, 14),
            At(request, "Zażółć", 42, 200, 58, 14)
        }, diagnostics: null, provider: "fixture", model: "fixture-v1", language: "eng"));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).Ocr.MakeSearchableAsync(provider);
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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "שלום", 150, 160, 48, 14),
            At(request, "עולם", 100, 160, 42, 14)
        }, diagnostics: null, provider: "fixture", model: "fixture-v1", language: "heb"));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).Ocr.MakeSearchableAsync(provider);
        byte[] searchable = result.Document.ToBytes();

        Assert.Contains("שלום עולם", result.Ocr.Pages[0].Text, StringComparison.Ordinal);
        Assert.Contains(result.Ocr.EnrichedDocument.TextBlocks,
            block => block.SourceKind == PdfLogicalContentSourceKind.Ocr && block.Text == "שלום עולם");
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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "LeftTop", 30, 80, 54, 12),
            At(request, "LeftBottom", 30, 120, 66, 12),
            At(request, "RightTop", 300, 80, 60, 12),
            At(request, "RightBottom", 300, 120, 72, 12)
        }));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).Ocr.MakeSearchableAsync(provider);
        string decodedContent = string.Join(
            Environment.NewLine,
            result.Document.Debug(new PdfDebuggerOptions {
                IncludeDecodedStreamPreviews = true,
                MaxDecodedStreamPreviewBytes = 64 * 1024
            }).Objects.Select(static item => item.DecodedStreamPreview ?? string.Empty));

        string[] expected = { "LeftTop", "LeftBottom", "RightTop", "RightBottom" };
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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "Once", 42, 160, 40, 14)
        }));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).Ocr.MakeSearchableAsync(provider, new PdfOcrMergeOptions {
            Selection = PdfPageSelection.From(1, 1),
            DetectAlignedTables = false
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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "Already searchable", 42, 160, 120, 14)
        }));

        PdfSearchableOcrResult first = await PdfDocument.Load(source).Ocr.MakeSearchableAsync(provider);
        PdfSearchableOcrResult second = await first.Document.Ocr.MakeSearchableAsync(provider);

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
        var provider = new StubOcrProvider(_ => new PdfOcrResponse(Array.Empty<PdfOcrWord>()));
        PdfDocument document = PdfDocument.Load(source);

        PdfSearchableOcrResult result = await document.Ocr.MakeSearchableAsync(provider);

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
        var provider = new StubOcrProvider(request => new PdfOcrResponse(new[] {
            At(request, "Rotated", 24, 180, 64, 12)
        }));

        PdfSearchableOcrResult result = await PdfDocument.Load(source).Ocr.MakeSearchableAsync(provider);
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

    private static PdfOcrWord At(PdfOcrRequest request, string text, double x, double y, double width, double height, double confidence = 0.95D) =>
        new PdfOcrWord(text, x * request.Scale, y * request.Scale, width * request.Scale, height * request.Scale, confidence);

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
