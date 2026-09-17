using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using OfficeIMO.Workflows;

namespace OfficeIMO.Workflows.Tests;

public sealed partial class RasterContentSafetyTests {
    [Fact]
    public async Task RedactionRejectsPaddingThatOverlapsUnselectedVisibleText() {
        var raster = new OfficeRasterImage(32, 16, OfficeColor.White);
        for (int y = 4; y < 10; y++) {
            for (int x = 3; x < 13; x++) raster.SetPixel(x, y, OfficeColor.FromRgb(248, 248, 248));
            for (int x = 16; x < 24; x++) raster.SetPixel(x, y, OfficeColor.Black);
        }
        byte[] image = OfficePngWriter.Encode(raster);
        IOcrEngine engine = CreateEngine(_ => new OcrResult {
            Text = "concealed visible",
            Spans = new[] {
                Span(0, "concealed", new OcrRegion { X = 3, Y = 4, Width = 10, Height = 6 }, 0.99D),
                Span(1, "visible", new OcrRegion { X = 16, Y = 4, Width = 8, Height = 6 }, 0.99D)
            }
        });
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            RedactionPaddingPixels = 4
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            OfficeRasterContentSafety.RedactSelectedContentAsync(
                image,
                engine,
                new OfficeContentCleanupSelection(new[] { finding.Id }),
                options));
    }

    [Fact]
    public async Task RedactionRejectsAnOverlappingConcealedSpanThatWasNotSelected() {
        byte[] image = CreateImage(40, 20, OfficeColor.White,
            new PixelBox(5, 6, 24, 6), OfficeColor.FromRgb(248, 248, 248));
        IOcrEngine engine = CreateEngine(_ => new OcrResult {
            Text = "concealed line",
            Spans = new[] {
                Span(0, "concealed", new OcrRegion { X = 5, Y = 6, Width = 12, Height = 6 }, 0.99D,
                    OcrTextSpanLevel.Word, "1:1:1:1"),
                Span(1, "concealed line", new OcrRegion { X = 5, Y = 6, Width = 24, Height = 6 }, 0.99D,
                    OcrTextSpanLevel.Line, "1:1:1:1")
            }
        });
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            RedactionPaddingPixels = 0
        };
        OfficeContentSafetyReport report = await OfficeRasterContentSafety.InspectAsync(image, engine, options);
        Assert.Equal(2, report.Findings.Count);

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            OfficeRasterContentSafety.RedactSelectedContentAsync(
                image,
                engine,
                new OfficeContentCleanupSelection(new[] { report.Findings[0].Id }),
                options));
    }

    [Fact]
    public async Task RedactionUsesFinerLineChildrenInsteadOfAggregateLineBounds() {
        var raster = new OfficeRasterImage(40, 20, OfficeColor.White);
        for (int y = 6; y < 12; y++) {
            for (int x = 2; x < 9; x++) raster.SetPixel(x, y, OfficeColor.Black);
            for (int x = 14; x < 21; x++) raster.SetPixel(x, y, OfficeColor.FromRgb(248, 248, 248));
            for (int x = 28; x < 35; x++) raster.SetPixel(x, y, OfficeColor.Black);
        }
        byte[] image = OfficePngWriter.Encode(raster);
        const string lineId = "1:1:1:1";
        int calls = 0;
        IOcrEngine engine = CreateEngine(_ => calls++ < 2
            ? new OcrResult {
                Text = "left concealed right",
                Spans = new[] {
                    Span(0, "left concealed right", new OcrRegion { X = 2, Y = 6, Width = 33, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Line, lineId),
                    Span(1, "left", new OcrRegion { X = 2, Y = 6, Width = 7, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Word, lineId),
                    Span(2, "concealed", new OcrRegion { X = 14, Y = 6, Width = 7, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Word, lineId),
                    Span(3, "right", new OcrRegion { X = 28, Y = 6, Width = 7, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Word, lineId)
                }
            }
            : new OcrResult {
                Text = "left right",
                Spans = new[] {
                    Span(0, "left right", new OcrRegion { X = 2, Y = 6, Width = 33, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Line, lineId),
                    Span(1, "left", new OcrRegion { X = 2, Y = 6, Width = 7, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Word, lineId),
                    Span(2, "right", new OcrRegion { X = 28, Y = 6, Width = 7, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Word, lineId)
                }
            });
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            RedactionPaddingPixels = 0
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        OfficeContentCleanupResult cleanup = await OfficeRasterContentSafety.RedactSelectedContentAsync(
            image,
            engine,
            new OfficeContentCleanupSelection(new[] { finding.Id }),
            options);

        Assert.True(cleanup.Changed);
        Assert.Single(cleanup.Changes);
    }

    [Fact]
    public async Task RedactionUsesCharacterChildrenInsteadOfAggregateWordBounds() {
        var raster = new OfficeRasterImage(30, 16, OfficeColor.White);
        for (int y = 5; y < 11; y++) {
            for (int x = 4; x < 12; x++) raster.SetPixel(x, y, OfficeColor.Black);
            for (int x = 14; x < 22; x++) raster.SetPixel(x, y, OfficeColor.FromRgb(248, 248, 248));
        }
        byte[] image = OfficePngWriter.Encode(raster);
        const string lineId = "1:1:1:1";
        int calls = 0;
        IOcrEngine engine = CreateEngine(_ => calls++ < 2
            ? new OcrResult {
                Text = "ab",
                Spans = new[] {
                    Span(0, "ab", new OcrRegion { X = 4, Y = 5, Width = 18, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Word, lineId),
                    Span(1, "a", new OcrRegion { X = 4, Y = 5, Width = 8, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Character, lineId),
                    Span(2, "b", new OcrRegion { X = 14, Y = 5, Width = 8, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Character, lineId)
                }
            }
            : new OcrResult {
                Text = "a",
                Spans = new[] {
                    Span(0, "a", new OcrRegion { X = 4, Y = 5, Width = 18, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Word, lineId),
                    Span(1, "a", new OcrRegion { X = 4, Y = 5, Width = 8, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Character, lineId)
                }
            });
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            RedactionPaddingPixels = 0
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        OfficeContentCleanupResult cleanup = await OfficeRasterContentSafety.RedactSelectedContentAsync(
            image,
            engine,
            new OfficeContentCleanupSelection(new[] { finding.Id }),
            options);

        Assert.True(cleanup.Changed);
        Assert.Single(cleanup.Changes);
    }

    [Theory]
    [InlineData(OcrDiagnosticSeverity.Error, true)]
    [InlineData(OcrDiagnosticSeverity.Warning, false)]
    public async Task RedactionRejectsFailedPostInspectionDiagnostics(
        OcrDiagnosticSeverity severity,
        bool isRecoverable) {
        byte[] image = CreateImage(30, 16, OfficeColor.White,
            new PixelBox(3, 4, 20, 6), OfficeColor.FromRgb(248, 248, 248));
        int calls = 0;
        IOcrEngine engine = CreateEngine(_ => calls++ < 2
            ? Result("concealed", new OcrRegion { X = 3, Y = 4, Width = 20, Height = 6 }, 0.99D)
            : new OcrResult {
                Diagnostics = new[] {
                    new OcrDiagnostic {
                        Severity = severity,
                        Code = "recognition-failed",
                        Message = "Recognition did not complete.",
                        IsRecoverable = isRecoverable
                    }
                }
            });
        var options = new OfficeRasterContentSafetyOptions { EnableOpaqueRectangleRedaction = true };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRasterContentSafety.RedactSelectedContentAsync(
                image,
                engine,
                new OfficeContentCleanupSelection(new[] { finding.Id }),
                options));
    }

    [Fact]
    public async Task RedactionRejectsPostInspectionAggregateTextOutsideBoundedSpans() {
        var raster = new OfficeRasterImage(40, 20, OfficeColor.White);
        for (int y = 5; y < 11; y++) {
            for (int x = 2; x < 10; x++) raster.SetPixel(x, y, OfficeColor.FromRgb(248, 248, 248));
            for (int x = 25; x < 33; x++) raster.SetPixel(x, y, OfficeColor.Black);
        }
        byte[] image = OfficePngWriter.Encode(raster);
        int calls = 0;
        IOcrEngine engine = CreateEngine(_ => calls++ < 2
            ? Result("concealed", new OcrRegion { X = 2, Y = 5, Width = 8, Height = 6 }, 0.99D)
            : new OcrResult {
                Text = "concealed visible",
                Spans = new[] {
                    Span(0, "visible", new OcrRegion { X = 25, Y = 5, Width = 8, Height = 6 }, 0.99D)
                }
            });
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            RedactionPaddingPixels = 0
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRasterContentSafety.RedactSelectedContentAsync(
                image,
                engine,
                new OfficeContentCleanupSelection(new[] { finding.Id }),
                options));
    }

    [Fact]
    public async Task RedactionBoundsRegionIntersectionComparisons() {
        var raster = new OfficeRasterImage(40, 20, OfficeColor.White);
        for (int y = 5; y < 11; y++) {
            for (int x = 2; x < 10; x++) raster.SetPixel(x, y, OfficeColor.FromRgb(248, 248, 248));
            for (int x = 20; x < 26; x++) raster.SetPixel(x, y, OfficeColor.Black);
            for (int x = 30; x < 36; x++) raster.SetPixel(x, y, OfficeColor.Black);
        }
        byte[] image = OfficePngWriter.Encode(raster);
        IOcrEngine engine = CreateEngine(_ => new OcrResult {
            Text = "concealed visible visible",
            Spans = new[] {
                Span(0, "concealed", new OcrRegion { X = 2, Y = 5, Width = 8, Height = 6 }, 0.99D),
                Span(1, "visible", new OcrRegion { X = 20, Y = 5, Width = 6, Height = 6 }, 0.99D),
                Span(2, "visible", new OcrRegion { X = 30, Y = 5, Width = 6, Height = 6 }, 0.99D)
            }
        });
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            MaximumRegionComparisons = 1,
            RedactionPaddingPixels = 0
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRasterContentSafety.RedactSelectedContentAsync(
                image,
                engine,
                new OfficeContentCleanupSelection(new[] { finding.Id }),
                options));
    }

    [Fact]
    public async Task RedactionSharesRegionComparisonBudgetWithPostInspection() {
        var raster = new OfficeRasterImage(40, 20, OfficeColor.White);
        for (int y = 5; y < 11; y++) {
            for (int x = 2; x < 10; x++) raster.SetPixel(x, y, OfficeColor.FromRgb(248, 248, 248));
            for (int x = 25; x < 33; x++) raster.SetPixel(x, y, OfficeColor.Black);
        }
        byte[] image = OfficePngWriter.Encode(raster);
        int calls = 0;
        IOcrEngine engine = CreateEngine(_ => calls++ < 2
            ? new OcrResult {
                Text = "concealed visible",
                Spans = new[] {
                    Span(0, "concealed", new OcrRegion { X = 2, Y = 5, Width = 8, Height = 6 }, 0.99D),
                    Span(1, "visible", new OcrRegion { X = 25, Y = 5, Width = 8, Height = 6 }, 0.99D)
                }
            }
            : Result("visible", new OcrRegion { X = 25, Y = 5, Width = 8, Height = 6 }, 0.99D));
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            MaximumRegionComparisons = 1,
            RedactionPaddingPixels = 0
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRasterContentSafety.RedactSelectedContentAsync(
                image,
                engine,
                new OfficeContentCleanupSelection(new[] { finding.Id }),
                options));
    }

    [Fact]
    public async Task RedactionSharesPixelWorkBudgetAcrossInspectionMutationAndVerification() {
        byte[] image = CreateImage(40, 20, OfficeColor.White,
            new PixelBox(2, 5, 8, 6), OfficeColor.FromRgb(248, 248, 248));
        int calls = 0;
        IOcrEngine engine = CreateEngine(_ => calls++ < 2
            ? Result("concealed", new OcrRegion { X = 2, Y = 5, Width = 8, Height = 6 }, 0.99D)
            : new OcrResult());
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            MaximumPixelAnalysisWork = 1_000,
            RedactionPaddingPixels = 0
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRasterContentSafety.RedactSelectedContentAsync(
                image,
                engine,
                new OfficeContentCleanupSelection(new[] { finding.Id }),
                options));
    }

    [Fact]
    public async Task RedactionRejectsASelectedRegionThatAlreadyMatchesTheRedactionColor() {
        byte[] image = CreateImage(20, 10, OfficeColor.White,
            new PixelBox(3, 4, 8, 3), OfficeColor.Black);
        IOcrEngine engine = CreateEngine(_ =>
            Result("tiny", new OcrRegion { X = 3, Y = 4, Width = 8, Height = 3 }, 0.99D));
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            RedactionPaddingPixels = 0,
            RedactionColor = OfficeColor.Black
        };
        OfficeContentSafetyFinding finding = Assert.Single(
            (await OfficeRasterContentSafety.InspectAsync(image, engine, options)).Findings);

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            OfficeRasterContentSafety.RedactSelectedContentAsync(
                image,
                engine,
                new OfficeContentCleanupSelection(new[] { finding.Id }),
                options));
    }

    [Fact]
    public async Task RedactionWithEveryRecognizedSpanSelectedNeedsNoPreComparisonBudget() {
        byte[] image = CreateImage(40, 20, OfficeColor.White,
            new PixelBox(5, 6, 24, 6), OfficeColor.FromRgb(248, 248, 248));
        const string lineId = "1:1:1:1";
        int calls = 0;
        IOcrEngine engine = CreateEngine(_ => calls++ < 2
            ? new OcrResult {
                Text = "concealed text",
                Spans = new[] {
                    Span(0, "concealed text", new OcrRegion { X = 5, Y = 6, Width = 24, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Line, lineId),
                    Span(1, "concealed", new OcrRegion { X = 5, Y = 6, Width = 11, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Word, lineId),
                    Span(2, "text", new OcrRegion { X = 18, Y = 6, Width = 11, Height = 6 }, 0.99D,
                        OcrTextSpanLevel.Word, lineId)
                }
            }
            : new OcrResult());
        var options = new OfficeRasterContentSafetyOptions {
            EnableOpaqueRectangleRedaction = true,
            MaximumRegionComparisons = 1,
            RedactionPaddingPixels = 0
        };
        OfficeContentSafetyReport report = await OfficeRasterContentSafety.InspectAsync(image, engine, options);
        Assert.Equal(3, report.Findings.Count);

        OfficeContentCleanupResult cleanup = await OfficeRasterContentSafety.RedactSelectedContentAsync(
            image,
            engine,
            new OfficeContentCleanupSelection(report.Findings.Select(finding => finding.Id).ToArray()),
            options);

        Assert.True(cleanup.Changed);
        Assert.Equal(3, cleanup.Changes.Count);
    }
}
