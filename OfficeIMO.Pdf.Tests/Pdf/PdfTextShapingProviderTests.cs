using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTextShapingProviderTests {
    [Fact]
    public void ManagedPdfUsesTheSharedTypographyCorpusAndRetainsLogicalText() {
        foreach (TypographyEvidenceCase evidence in TypographyEvidenceCorpus.Cases) {
            byte[] fontData = LoadTypographyFont(evidence);
            var report = new PdfConversionReport();
            var options = new PdfOptions { CompressContentStreams = false }
                .ReportDiagnosticsTo(report, "OfficeIMO.Pdf.Tests")
                .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(evidence.Family, fontData))
                .SetLanguage(evidence.Language)
                .SetTextShapingProvider(OfficeManagedTextShapingProvider.Instance);

            byte[] bytes = PdfDocument.Create(options)
                .Paragraph(paragraph => paragraph.FontFamily(evidence.Family).Text(evidence.Text))
                .ToBytes();

            string extracted = PdfReadDocument.Open(bytes).ExtractText();
            foreach (string logicalToken in evidence.Text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                Assert.Contains(logicalToken, extracted, StringComparison.Ordinal);
            }
            if (evidence.Direction == OfficeTextDirection.RightToLeft || evidence.ManagedShapingExpected) {
                Assert.Contains("/ActualText", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
            }
            Assert.DoesNotContain(report.Warnings, warning =>
                warning.Code.Contains("font-family-substitution", StringComparison.OrdinalIgnoreCase));
        }
    }

    [Fact]
    public void VerticalDrawingRetainsLogicalTextAndReportsThePdfPositioningFallback() {
        TypographyEvidenceCase evidence = Assert.Single(
            TypographyEvidenceCorpus.Cases,
            item => item.Direction == OfficeTextDirection.TopToBottom);
        byte[] fontData = LoadTypographyFont(evidence);
        var drawing = new OfficeDrawing(120D, 180D)
            .AddFont(evidence.Family, fontData)
            .AddVerticalText(evidence.Text, 20D, 10D, 80D, 160D, new OfficeFontInfo(evidence.Family, 36D));
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false,
                PageWidth = 180D,
                PageHeight = 240D,
                MarginLeft = 20D,
                MarginRight = 20D,
                MarginTop = 20D,
                MarginBottom = 20D
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Pdf.Tests")
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(evidence.Family, fontData))
            .SetLanguage(evidence.Language)
            .SetTextShapingProvider(OfficeManagedTextShapingProvider.Instance);

        byte[] bytes = PdfDocument.Create(options).Drawing(drawing).ToBytes();
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains(evidence.Text, extracted, StringComparison.Ordinal);
        Assert.Single(report.FidelityDiagnostics, diagnostic =>
            diagnostic.Code == "vertical-text-stacked-fallback" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());
        Assert.Contains("/ActualText", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void VerticalDrawingInsideActualTextGroupStillReportsThePdfPositioningFallback() {
        TypographyEvidenceCase evidence = Assert.Single(
            TypographyEvidenceCorpus.Cases,
            item => item.Direction == OfficeTextDirection.TopToBottom);
        byte[] fontData = LoadTypographyFont(evidence);
        var paint = new OfficeDrawing(100D, 160D)
            .AddFont(evidence.Family, fontData)
            .AddVerticalText(evidence.Text, 10D, 10D, 80D, 140D, new OfficeFontInfo(evidence.Family, 32D));
        var drawing = new OfficeDrawing(100D, 160D)
            .AddActualTextDrawing(evidence.Text, paint, 50D, 20D);
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false,
                PageWidth = 140D,
                PageHeight = 200D,
                MarginLeft = 20D,
                MarginRight = 20D,
                MarginTop = 20D,
                MarginBottom = 20D
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Pdf.Tests")
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(evidence.Family, fontData));

        byte[] bytes = PdfDocument.Create(options).Drawing(drawing).ToBytes();

        Assert.Single(report.FidelityDiagnostics, diagnostic =>
            diagnostic.Code == "vertical-text-stacked-fallback" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());
        Assert.Contains(evidence.Text, PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void ImportedAuthoredDirectionReachesPdfShapingAndRetainsLogicalText() {
        const string value = "abc אבג";
        const string family = "Direction Contract";
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 320 60'>"
            + "<text x='300' y='30' font-family='Direction Contract' font-size='18' fill='black' "
            + "direction='rtl' text-anchor='start'>abc אבג</text></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg), options: null, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);

        byte[] fontData = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
            'a', 'b', 'c', ' ', 0x05D0, 0x05D1, 0x05D2);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, family);
        var glyphs = new List<OfficeShapedGlyph>();
        for (int index = 0; index < value.Length; index++) {
            Assert.True(fontProgram.TryGetGlyphId(value[index], out int glyphId));
            glyphs.Add(new OfficeShapedGlyph(glyphId, value[index].ToString(), index, fontProgram.UnitsPerEm / 2));
        }
        var provider = new DirectionRecordingTextShapingProvider(value, glyphs);
        var options = new PdfOptions { CompressContentStreams = false }
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, fontData))
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options).Drawing(drawing!).ToBytes();

        Assert.Contains(provider.Requests, request =>
            request.Direction == OfficeTextDirection.RightToLeft && !request.IsOpenTypeCff);
        Assert.Contains(value, PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("/ActualText", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void HorizontalCorpusHasComparableRenderedGeometryAcrossRasterSvgAndPdf() {
        foreach (TypographyEvidenceCase evidence in TypographyEvidenceCorpus.Cases.Where(
            item => item.Direction != OfficeTextDirection.TopToBottom)) {
            byte[] fontData = LoadTypographyFont(evidence);
            var drawing = new OfficeDrawing(560D, 100D)
                .AddFont(evidence.Family, fontData)
                .AddText(evidence.Text, 10D, 10D, 540D, 80D, new OfficeFontInfo(evidence.Family, 28D));
            var rasterOptions = new OfficeDrawingRasterRenderOptions {
                TextShapingProvider = OfficeManagedTextShapingProvider.Instance,
                TextShapingLanguage = evidence.Language
            };
            drawing.ApplyImageExportOptions(new OfficeImageExportOptions {
                TextShapingProvider = OfficeManagedTextShapingProvider.Instance,
                TextShapingLanguage = evidence.Language
            });
            PixelBounds directBounds = GetInkBounds(OfficeDrawingRasterRenderer.Render(drawing, rasterOptions));

            byte[] svg = OfficeDrawingSvgExporter.ToSvgBytes(drawing, 1D, OfficeSvgSizeUnit.Pixel);
            var svgOptions = new OfficeSvgDrawingReaderOptions();
            svgOptions.Fonts.Add(evidence.Family, fontData);
            Assert.True(OfficeSvgDrawingReader.TryRead(svg, svgOptions, out OfficeDrawing? svgDrawing), evidence.Name);
            PixelBounds svgBounds = GetInkBounds(OfficeDrawingRasterRenderer.Render(svgDrawing!, rasterOptions));

            var pdfOptions = new PdfOptions {
                    PageWidth = 600D,
                    PageHeight = 140D,
                    MarginLeft = 0D,
                    MarginRight = 0D,
                    MarginTop = 0D,
                    MarginBottom = 0D
                }
                .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(evidence.Family, fontData))
                .SetLanguage(evidence.Language)
                .SetTextShapingProvider(OfficeManagedTextShapingProvider.Instance);
            byte[] pdf = PdfDocument.Create(pdfOptions)
                .Canvas(canvas => canvas.Drawing(drawing, 20D, 20D, 560D, 100D))
                .ToBytes();
            PixelBounds pdfBounds = GetInkBounds(OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf), rasterOptions));

            AssertGeometryClose(directBounds, svgBounds, evidence.Name + " raster/SVG", offsetX: 0, offsetY: 0);
            AssertGeometryClose(directBounds, pdfBounds, evidence.Name + " raster/PDF", offsetX: 20, offsetY: 20);
            foreach (string token in evidence.Text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                Assert.Contains(token, PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
            }
        }
    }

    private static PixelBounds GetInkBounds(OfficeRasterImage image) {
        int minX = image.Width, minY = image.Height, maxX = -1, maxY = -1;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.A == 0 || pixel.R >= 250 && pixel.G >= 250 && pixel.B >= 250) continue;
                minX = Math.Min(minX, x); minY = Math.Min(minY, y);
                maxX = Math.Max(maxX, x); maxY = Math.Max(maxY, y);
            }
        }
        Assert.True(maxX >= minX && maxY >= minY, "Expected rendered text ink.");
        return new PixelBounds(minX, minY, maxX - minX + 1, maxY - minY + 1);
    }

    private static void AssertGeometryClose(PixelBounds expected, PixelBounds actual, string source, int offsetX, int offsetY) {
        int positionTolerance = Math.Max(8, (int)Math.Ceiling(expected.Width * 0.25D));
        int widthTolerance = Math.Max(6, (int)Math.Ceiling(expected.Width * 0.35D));
        int heightTolerance = Math.Max(5, (int)Math.Ceiling(expected.Height * 0.35D));
        PixelBounds normalized = actual with { X = actual.X - offsetX, Y = actual.Y - offsetY };
        Assert.True(normalized.X >= expected.X - positionTolerance && normalized.X <= expected.X + positionTolerance,
            $"{source} X mismatch: expected {expected}, actual {normalized}.");
        Assert.True(normalized.Y >= expected.Y - 10 && normalized.Y <= expected.Y + 10,
            $"{source} Y mismatch: expected {expected}, actual {normalized}.");
        Assert.True(normalized.Width >= expected.Width - widthTolerance && normalized.Width <= expected.Width + widthTolerance,
            $"{source} width mismatch: expected {expected}, actual {normalized}.");
        Assert.True(normalized.Height >= expected.Height - heightTolerance && normalized.Height <= expected.Height + heightTolerance,
            $"{source} height mismatch: expected {expected}, actual {normalized}.");
    }

    private readonly record struct PixelBounds(int X, int Y, int Width, int Height);

    [Fact]
    public void TextShapingProvider_ShapesEmbeddedTrueTypeComplexScriptWithoutUnsupportedWarnings() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string text = "\u0633\u0644\u0627\u0645";
        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Provider Font");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
            return;
        }

        var provider = new MappingTextShapingProvider(text, isOpenTypeCff: false, CreateGlyphMap(text, fontProgram));
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Provider Font")
            .SetLanguage("ar-SA")
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.True(provider.CallCount >= 1);
        Assert.NotNull(provider.LastRequest);
        Assert.Equal(OfficeTextDirection.RightToLeft, provider.LastRequest!.Direction);
        Assert.Equal("ar-SA", provider.LastRequest.Language);
        Assert.Equal(fontProgram.UnitsPerEm, provider.LastRequest.UnitsPerEm);
        Assert.Contains(text, extracted, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
    }

    [Fact]
    public void TextShapingProvider_DoesNotSuppressWarningsWhenProviderDeclinesRun() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string text = "\u0633\u0644\u0627\u0645";
        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Provider Font");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
            return;
        }

        var provider = new DecliningTextShapingProvider();
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Provider Font")
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        string raw = Encoding.ASCII.GetString(bytes);

        Assert.True(provider.CallCount >= 1);
        Assert.Contains("/ActualText", raw, StringComparison.Ordinal);
        Assert.Contains(text, extracted, StringComparison.Ordinal);
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
    }

    [Fact]
    public void TextShapingProvider_SuppressesWarningsForAutomaticallyPlannedFallbackRun() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string text = "\u0633\u0644\u0627\u0645";
        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Provider Fallback");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("OfficeIMO Provider Fallback", fontData) },
            new[] { PdfStandardFont.TimesRoman });
        var provider = new MappingTextShapingProvider(text, isOpenTypeCff: false, CreateGlyphMap(text, fontProgram));
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.True(provider.CallCount >= 1);
        Assert.Contains(text, extracted, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
    }

    [Fact]
    public void TextShapingProvider_DoesNotSuppressFallbackWarningsWhenProviderDeclinesRun() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string text = "\u0633\u0644\u0627\u0645";
        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Provider Fallback");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, fontProgram).Count > 0) {
            return;
        }

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("OfficeIMO Provider Fallback", fontData) },
            new[] { PdfStandardFont.TimesRoman });
        var provider = new DecliningTextShapingProvider();
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .RegisterEmbeddedFontFallbacks(fallbackSet)
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.True(provider.CallCount >= 1);
        Assert.Contains(text, extracted, StringComparison.Ordinal);
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
        Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
    }

    [Fact]
    public void TextShapingProvider_MapsOpenTypeCffLigatureGlyphBackToSourceText() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] fontData = File.ReadAllBytes(fontPath!);
        PdfOpenTypeCffFontProgram fontProgram = PdfOpenTypeCffFontProgram.Parse(fontData, "OfficeIMO Source Serif CFF");
        Assert.True(fontProgram.TryGetGlyphId('o', out int oGlyphId));
        Assert.True(fontProgram.TryGetGlyphId(0xFB03, out int ffiGlyphId));
        Assert.True(fontProgram.TryGetGlyphId('c', out int cGlyphId));
        Assert.True(fontProgram.TryGetGlyphId('e', out int eGlyphId));

        var provider = new MappingTextShapingProvider(
            "office",
            isOpenTypeCff: true,
            new[] {
                new OfficeShapedGlyph(oGlyphId, "o", 0),
                new OfficeShapedGlyph(ffiGlyphId, "ffi", 1),
                new OfficeShapedGlyph(cGlyphId, "c", 4),
                new OfficeShapedGlyph(eGlyphId, "e", 5)
            });
        var options = new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif CFF")
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("office"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.True(provider.CallCount >= 1);
        Assert.Contains("office", extracted, StringComparison.Ordinal);
        Assert.Contains("<" + ffiGlyphId.ToString("X4", CultureInfo.InvariantCulture) + "> <006600660069>", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void TextShapingProvider_WritesDesignUnitAdvancesOffsetsAndLogicalActualText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMO Positioned Provider Font");
        Assert.True(fontProgram.TryGetGlyphId('A', out int aGlyphId));
        Assert.True(fontProgram.TryGetGlyphId('B', out int bGlyphId));
        int halfEm = fontProgram.UnitsPerEm / 2;
        int offsetX = -(fontProgram.UnitsPerEm / 10);
        int offsetY = fontProgram.UnitsPerEm / 5;
        var provider = new MappingTextShapingProvider(
            "AB",
            isOpenTypeCff: false,
            new[] {
                new OfficeShapedGlyph(aGlyphId, "A", 0, halfEm),
                new OfficeShapedGlyph(bGlyphId, "B", 1, halfEm, offsetX, offsetY)
            });
        var options = new PdfOptions {
                CompressContentStreams = false
            }
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Positioned Provider Font")
            .SetLanguage("en-US")
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("AB"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        double measured = fontProgram.MeasureTextWidth("AB", 12D, shapingProvider: provider, language: "en-US");

        Assert.Equal(12D * (halfEm + halfEm) / fontProgram.UnitsPerEm, measured, precision: 6);
        Assert.Contains("AB", extracted, StringComparison.Ordinal);
        Assert.Contains(" TJ", raw, StringComparison.Ordinal);
        Assert.Contains(" Ts", raw, StringComparison.Ordinal);
        Assert.Contains("/ActualText", raw, StringComparison.Ordinal);
        Assert.NotNull(provider.LastRequest);
        Assert.Equal(OfficeTextDirection.LeftToRight, provider.LastRequest!.Direction);
        Assert.Equal("en-US", provider.LastRequest.Language);
        Assert.Equal(fontProgram.UnitsPerEm, provider.LastRequest.UnitsPerEm);
    }

    [Fact]
    public void TextShapingProvider_PreservesSuperscriptRiseAroundPositionedGlyphs() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] fontData = File.ReadAllBytes(fontPath!);
        PdfOpenTypeCffFontProgram fontProgram = PdfOpenTypeCffFontProgram.Parse(fontData, "OfficeIMO Positioned Superscript Font");
        Assert.True(fontProgram.TryGetGlyphId('A', out int aGlyphId));
        Assert.True(fontProgram.TryGetGlyphId('B', out int bGlyphId));
        int offsetY = fontProgram.UnitsPerEm / 5;
        var provider = new MappingTextShapingProvider(
            "AB",
            isOpenTypeCff: true,
            new[] {
                new OfficeShapedGlyph(aGlyphId, "A", 0, fontProgram.UnitsPerEm),
                new OfficeShapedGlyph(bGlyphId, "B", 1, fontProgram.UnitsPerEm, 0, offsetY)
            });
        var options = new PdfOptions {
                CompressContentStreams = false,
                DefaultFontSize = 12D
            }
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Positioned Superscript Font")
            .SetTextShapingProvider(provider);

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Superscript("AB").Text("C"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        const double baseRise = 12D * 0.35D;
        const double runFontSize = 12D * 0.65D;
        int offsetY1000 = checked((int)Math.Round(offsetY * 1000D / fontProgram.UnitsPerEm, MidpointRounding.AwayFromZero));
        double positionedRise = baseRise + offsetY1000 * runFontSize / 1000D;
        string baseRiseOperator = baseRise.ToString("0.###", CultureInfo.InvariantCulture) + " Ts";
        string positionedRiseOperator = positionedRise.ToString("0.###", CultureInfo.InvariantCulture) + " Ts";
        int positionedRiseIndex = raw.IndexOf(positionedRiseOperator, StringComparison.Ordinal);
        int restoredBaseRiseIndex = raw.IndexOf(baseRiseOperator, positionedRiseIndex + positionedRiseOperator.Length, StringComparison.Ordinal);
        int normalRiseIndex = raw.IndexOf("0 Ts", restoredBaseRiseIndex + baseRiseOperator.Length, StringComparison.Ordinal);

        Assert.True(positionedRiseIndex >= 0, "Expected the shaped glyph offset to be added to the superscript rise.");
        Assert.True(restoredBaseRiseIndex > positionedRiseIndex, "Expected the superscript rise to be restored after positioned glyphs.");
        Assert.True(normalRiseIndex > restoredBaseRiseIndex, "Expected normal text to reset the restored superscript rise.");
    }

    private static IReadOnlyList<OfficeShapedGlyph> CreateGlyphMap(string text, PdfTrueTypeFontProgram fontProgram) {
        var glyphs = new List<OfficeShapedGlyph>();
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            int scalar = ReadScalar(text, ref index);
            Assert.True(fontProgram.TryGetGlyphId(scalar, out int glyphId));
            glyphs.Add(new OfficeShapedGlyph(glyphId, char.ConvertFromUtf32(scalar), scalarStart));
        }

        return glyphs;
    }

    private static byte[] LoadTypographyFont(TypographyEvidenceCase evidence) {
        if (!string.IsNullOrEmpty(evidence.FontFileName)) {
            return File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", evidence.FontFileName));
        }
        return evidence.Name == "Hebrew"
            ? ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x05E9, 0x05DC, 0x05D5, 0x05DD, 0x05E2)
            : ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'C', 'a', 'f', 'e', 0x0301);
    }

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch) && index < text.Length && char.IsLowSurrogate(text[index])) {
            return char.ConvertToUtf32(ch, text[index++]);
        }

        return ch;
    }

    private sealed class MappingTextShapingProvider : IOfficeTextShapingProvider {
        private readonly string _text;
        private readonly bool _isOpenTypeCff;
        private readonly IReadOnlyList<OfficeShapedGlyph> _glyphs;

        public MappingTextShapingProvider(string text, bool isOpenTypeCff, IReadOnlyList<OfficeShapedGlyph> glyphs) {
            _text = text;
            _isOpenTypeCff = isOpenTypeCff;
            _glyphs = glyphs;
        }

        public int CallCount { get; private set; }

        public OfficeTextShapingRequest? LastRequest { get; private set; }

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            if (!string.Equals(request.Text, _text, StringComparison.Ordinal)) {
                return null;
            }

            Assert.Equal(_isOpenTypeCff, request.IsOpenTypeCff);
            Assert.NotEmpty(request.FontData);
            Assert.False(string.IsNullOrWhiteSpace(request.FontName));
            CallCount++;
            LastRequest = request;
            return new OfficeTextShapingResult(_glyphs);
        }
    }

    private sealed class DirectionRecordingTextShapingProvider : IOfficeTextShapingProvider {
        private readonly string _text;
        private readonly IReadOnlyList<OfficeShapedGlyph> _glyphs;

        internal DirectionRecordingTextShapingProvider(string text, IReadOnlyList<OfficeShapedGlyph> glyphs) {
            _text = text;
            _glyphs = glyphs;
        }

        internal List<OfficeTextShapingRequest> Requests { get; } = new();

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            if (!string.Equals(request.Text, _text, StringComparison.Ordinal)) return null;
            Requests.Add(request);
            return new OfficeTextShapingResult(_glyphs, request.Direction);
        }
    }

    private sealed class DecliningTextShapingProvider : IOfficeTextShapingProvider {
        public int CallCount { get; private set; }

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            CallCount++;
            return null;
        }
    }
}
