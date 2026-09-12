using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void ToDrawing_WithFontProfileHonorsCancellationBeforeShaping() {
        byte[] pdf = BuildSingleStreamPdf("BT /F1 10 Tf 1 Tc 10 100 Td (AB) Tj ET");
        var provider = new OfficeIMO.TestAssets.ManagedTextShapingTestAssets.RecordingProvider();
        var fonts = new OfficeFontFaceCollection();
        fonts.Add("Courier New", OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont('A', 'B'));
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => PdfReadDocument.Open(pdf).Pages[0]
            .ToDrawing(fonts, provider, cancellationToken: cancellation.Token));
        Assert.Empty(provider.Requests);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    public void ExportImage_MultiCharacterGlyphMappingsChargeMaterializationBeforeMeasurement(int glyphCount) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier /ToUnicode 6 0 R >>\nendobj";
        const string cmap = "1 beginbfchar\n<41> <00410042>\nendbfchar";
        byte[] pdf = BuildSingleStreamPdf("BT /F1 10 Tf 1 Tc 1000 100 Td (" + new string('A', glyphCount) + ") Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font, BuildStreamObject(6, "<<", cmap));
        var document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextWorkCharactersPerPage = glyphCount * 2 - 1 }
        });
        var provider = new OfficeIMO.TestAssets.ManagedTextShapingTestAssets.RecordingProvider();
        var options = new PdfImageExportOptions { TextShapingProvider = provider };
        options.Fonts.Add("Courier New", OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont('A', 'B'));
        var error = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ExportImage(OfficeImageExportFormat.Png, options));
        Assert.Equal(PdfReadLimitKind.PositionedTextWorkCharacters, error.Kind);
        Assert.Equal(glyphCount * 2, error.Actual);
        if (glyphCount == 1) Assert.Empty(provider.Requests);
        else Assert.NotEmpty(provider.Requests);
    }

    [Fact]
    public void ExportImage_SpacesWithVisibleCustomFontInkStillConsumePositioningBudget() {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /F1 10 Tf -6 Tc 20 100 Td (A B) Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        var document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = 2 }
        });
        var options = new PdfImageExportOptions();
        options.Fonts.Add("Courier New", OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B'));
        var error = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ExportImage(OfficeImageExportFormat.Png, options));
        Assert.Equal(PdfReadLimitKind.PositionedTextCharacters, error.Kind);
        Assert.Equal(3, error.Actual);
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    public void ExportImage_BoundsInvisibleGlyphWorkBeforeCallingShapingProvider(bool forms, int spacing) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        const string resources = "<< /Font << /F1 5 0 R >> >>";
        string content = $"BT /F1 10 Tf {spacing} Tc 1000 100 Td " + (forms ? "(A) Tj" : "(A) Tj (B) Tj (C) Tj") + " ET";
        byte[] pdf = forms
            ? BuildSingleStreamPdf("/Fm Do /Fm Do /Fm Do", "<< /XObject << /Fm 6 0 R >> >>", font,
                BuildStreamObject(6, "<< /Type /XObject /Subtype /Form /BBox [0 0 2000 200] /Resources " + resources, content))
            : BuildSingleStreamPdf(content, resources, font);
        var document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextWorkCharactersPerPage = 2 }
        });
        var provider = new OfficeIMO.TestAssets.ManagedTextShapingTestAssets.RecordingProvider();
        var options = new PdfImageExportOptions { TextShapingProvider = provider };
        options.Fonts.Add("Courier New", OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont('A', 'B', 'C'));
        var error = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ExportImage(OfficeImageExportFormat.Png, options));
        Assert.Equal(PdfReadLimitKind.PositionedTextWorkCharacters, error.Kind);
        Assert.Equal(2, error.Limit);
        Assert.Equal(3, error.Actual);
        Assert.NotEmpty(provider.Requests);
        Assert.DoesNotContain(provider.Requests, request => request.Text.Contains("C"));
    }

    [Fact]
    public void RenderPage_RepeatedInvisibleGlyphsReuseOneWorkCharge() {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /F1 10 Tf 1 Tc 1000 100 Td (" + new string('A', 100001) + ") Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        var document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextWorkCharactersPerPage = 1, MaxPositionedTextCharactersPerPage = 1 }
        });
        Assert.Empty(document.Pages[0].ToDrawing().Elements);
        Assert.Empty(document.Pages[0].ToDrawing().Elements);
    }

    [Fact]
    public void RenderPage_OverprintedBlankSpacesDoNotConsumePositioningBudget() {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        string content = string.Concat(Enumerable.Repeat("A ", 50001)) + "B";
        byte[] pdf = BuildSingleStreamPdf("BT /F1 10 Tf -6 Tc 20 100 Td (" + content + ") Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        var drawing = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextWorkCharactersPerPage = 3 }
        }).Pages[0].ToDrawing();
        Assert.Equal(50002, drawing.Elements.Count);
        Assert.All(drawing.Elements.Take(50001), element => Assert.Equal("A", Assert.IsType<OfficeDrawingText>(element).Text));
        Assert.Equal("B", Assert.IsType<OfficeDrawingText>(drawing.Elements.Last()).Text);
    }

    [Theory]
    [InlineData(0, "page")]
    [InlineData(1, "page")]
    [InlineData(0, "blend")]
    [InlineData(1, "blend")]
    [InlineData(0, "pattern")]
    [InlineData(1, "pattern")]
    [InlineData(0, "mask")]
    [InlineData(1, "mask")]
    [InlineData(1, "shaped-page")]
    [InlineData(1, "shaped-blend")]
    [InlineData(1, "shaped-pattern")]
    [InlineData(1, "shaped-mask")]
    public void ExportImage_UsesCallerFontBeforeCullingEdgeGlyph(int spacing, string host) {
        bool shaped = host.StartsWith("shaped-", StringComparison.Ordinal);
        if (shaped) host = host.Substring("shaped-".Length);
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        string text = $"BT /F1 10 Tf {spacing} Tc 0 1 -1 0 260 100 Tm (A) Tj ET";
        byte[] pdf = BuildSingleStreamPdf((host == "blend" ? "/GS gs " : "") + text,
            "<< /Font << /F1 5 0 R >> /ExtGState << /GS 6 0 R >> >>", font,
            "6 0 obj\n<< /Type /ExtGState /BM /Multiply >>\nendobj");
        if (host == "pattern") {
            pdf = BuildSingleStreamPdf("/Pattern cs /P scn 0 0 240 200 re f", "<< /Pattern << /P 6 0 R >> >>", font,
                "6 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 240 200] /XStep 240 /YStep 200 " +
                $"/Resources << /Font << /F1 5 0 R >> >> /Length {text.Length} >>\nstream\n{text}\nendstream\nendobj");
        } else if (host == "mask") {
            pdf = BuildSingleStreamPdf("/GS gs 0 0 240 200 re f", "<< /ExtGState << /GS 6 0 R >> >>", font,
                "6 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 7 0 R >> >>\nendobj",
                "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Group << /S /Transparency >> " +
                $"/Resources << /Font << /F1 5 0 R >> >> /Length {text.Length} >>\nstream\n{text}\nendstream\nendobj");
        }
        var options = new PdfImageExportOptions { BackgroundColor = OfficeColor.Transparent };
        Assert.True(options.Fonts.TryAdd("Courier New", OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithTallGlyph('A', shaped ? 700 : 4000)));
        if (shaped) {
            options.TextShapingProvider = new RaisedEdgeGlyphProvider();
            options.TextShapingLanguage = "pl-PL";
        }
        var expected = new OfficeDrawing(240, 200).ApplyImageExportOptions(options);
        expected.AddClippedPositionedText("A", 260, 90, 6, 12.5, 0, 0, OfficeClipPath.Rectangle(240, 200),
            new OfficeImageFrameTransform(-90, 260, 100), new OfficeFontInfo("Courier New", 10),
            OfficeColor.Black, textAdvanceWidth: 6);
        byte[] pixels = OfficeDrawingRasterRenderer.Render(expected).GetPixels();
        Assert.Contains(pixels, channel => channel != 0);
        var projected = PdfReadDocument.Open(pdf).Pages[0].ToDrawing(options.Fonts,
            options.TextShapingProvider, options.TextShapingLanguage);
        Assert.Equal(pixels, OfficeDrawingRasterRenderer.Render(projected).GetPixels());
        var exported = PdfReadDocument.Open(pdf).Pages[0].ExportImage(OfficeImageExportFormat.Png, options);
        Assert.True(OfficePngReader.TryDecode(exported.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(pixels, raster!.GetPixels());
        var renderOptions = new PdfPageRenderOptions {
            Fonts = options.Fonts, Background = OfficeColor.Transparent,
            TextShapingProvider = options.TextShapingProvider, TextShapingLanguage = options.TextShapingLanguage
        };
        foreach (bool forDisplay in new[] { false, true }) {
            var rendered = PdfPageImageRenderer.RenderPage(PdfReadDocument.Open(pdf), 1, renderOptions,
                System.Threading.CancellationToken.None, forDisplay);
            Assert.True(OfficePngReader.TryDecode(rendered.Bytes!, out OfficeRasterImage? batchRaster));
            Assert.Equal(pixels, batchRaster!.GetPixels());
        }
    }

    private sealed class RaisedEdgeGlyphProvider : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Assert.Equal("pl-PL", request.Language);
            return new OfficeTextShapingResult(new[] { new OfficeShapedGlyph(1, request.Text, 0, 500, offsetY: 3000) });
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RenderPage_InvisibleLongOrdinaryTextDoesNotMeasureOutlines(bool emptyClip) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        string prefix = emptyClip ? "0 0 1 1 re W n 10 10 1 1 re W n BT /F1 10 Tf 20 100 Td (" : "BT /F1 10 Tf 1000 100 Td (";
        byte[] pdf = BuildSingleStreamPdf(prefix + new string('A', 1000000) + ") Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        Assert.Empty(PdfReadDocument.Open(pdf).Pages[0].ToDrawing().Elements);
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(true, 1)]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    public void RenderPage_RotatedInkCrossesEdgeWhenOriginIsOutside(bool clipped, int spacing) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        double origin = clipped ? 141 : 241;
        string clip = clipped ? "110 80 30 40 re W n " : "";
        byte[] pdf = BuildSingleStreamPdf(clip + $"BT /F1 10 Tf {spacing} Tc 0 1 -1 0 {origin} 100 Tm (A) Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        OfficeDrawing actual = PdfPageImageRenderer.RenderPage(pdf);
        var expected = new OfficeDrawing(240, 200);
        expected.AddClippedPositionedText("A", origin, 90, 6, 12.5,
            clipped ? 110 : 0, clipped ? 80 : 0,
            OfficeClipPath.Rectangle(clipped ? 30 : 240, clipped ? 40 : 200),
            new OfficeImageFrameTransform(-90, origin, 100), new OfficeFontInfo("Courier New", 10),
            OfficeColor.Black, textAdvanceWidth: 6);
        byte[] pixels = OfficeDrawingRasterRenderer.Render(expected).GetPixels();
        Assert.Contains(pixels, channel => channel != 0);
        Assert.Equal(pixels, OfficeDrawingRasterRenderer.Render(actual).GetPixels());
    }

    [Theory]
    [InlineData(-1, 100, "0 -1 1 0", 270)]
    [InlineData(100, 201, "-1 0 0 -1", 180)]
    [InlineData(100, -1, "1 0 0 1", 0)]
    public void RenderPage_PreservesTextOriginsAtOtherPageEdges(int x, int pdfY, string matrix, int rotation) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf($"BT /F1 10 Tf 1 Tc {matrix} {x} {pdfY} Tm (A) Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        var expected = new OfficeDrawing(240, 200);
        expected.AddClippedPositionedText("A", x, 200 - pdfY - 10, 6, 12.5, 0, 0, OfficeClipPath.Rectangle(240, 200),
            new OfficeImageFrameTransform(-rotation, x, 200 - pdfY), new OfficeFontInfo("Courier New", 10),
            OfficeColor.Black, textAdvanceWidth: 6);
        byte[] pixels = OfficeDrawingRasterRenderer.Render(expected).GetPixels();
        Assert.Contains(pixels, channel => channel != 0);
        Assert.Equal(pixels, OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf)).GetPixels());
    }

    [Fact]
    public void RenderPage_SpacedRunRetainsComplexClipOnce() {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        const string resources = "<< /Font << /F1 5 0 R >> >>";
        var clip = new StringBuilder("10 80 m ");
        for (int index = 0; index < 128; index++) clip.Append("60 80 l 60 120 l 10 120 l 10 80 l ");
        clip.Append("h W n ");
        const int glyphCount = 64;
        byte[] pdf = BuildSingleStreamPdf(clip + "BT /F1 10 Tf -6 Tc 20 100 Td (" + new string('A', glyphCount) + ") Tj ET", resources, font);
        OfficeDrawing actual = PdfPageImageRenderer.RenderPage(pdf);
        int retainedCommands = CountRetainedClipCommands(actual);
        // Retained scene work must grow with path size plus glyph count, not their product.
        Assert.True(retainedCommands <= 1024, "Retained clip commands: " + retainedCommands);
        var reference = new StringBuilder(clip.ToString());
        for (int index = 0; index < glyphCount; index++) reference.Append("BT /F1 10 Tf 20 100 Td (A) Tj ET ");
        OfficeDrawing expected = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(reference.ToString(), resources, font));
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected).GetPixels(), OfficeDrawingRasterRenderer.Render(actual).GetPixels());
    }

    private static int CountRetainedClipCommands(OfficeDrawing drawing) => drawing.Elements.OfType<OfficeDrawingGroup>()
        .Sum(group => group.ClipPath.Commands.Count + CountRetainedClipCommands(group.Drawing));

    [Theory]
    [InlineData(-6)]
    [InlineData(0)]
    public void RenderPage_TextWithEmptyFittedPathDoesNotPaint(int spacing) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        // The bounding box overlaps the page, but the triangle only touches its top-left corner.
        byte[] pdf = BuildSingleStreamPdf("-100 300 m -100 100 l 100 300 l h W n BT /F1 10 Tf " + spacing + " Tc 20 180 Td (AAAA) Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = 1 }
        });
        Assert.Empty(document.Pages[0].ToDrawing().Elements);
    }

    [Theory]
    [InlineData("90 80 m 120 80 l 120 120 l h W n ", false)]
    [InlineData("80 80 m 120 80 l 80 140 l h W n ", true)]
    [InlineData("90 80 40 40 re 108 95 5 10 re W* n ", false)]
    [InlineData("90 80 40 40 re 98 105 10 5 re W* n ", true)]
    [InlineData("-10 80 m 120 80 l 120 120 l -10 120 l h W n ", false)]
    [InlineData("90 80 m 140 80 140 120 90 120 c h W n ", false)]
    public void RenderPage_ComplexClippedSpacedTextMatchesExplicitOrigins(string clip, bool rotated) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        const string resources = "<< /Font << /F1 5 0 R >> >>";
        string matrix = rotated ? "0 1 -1 0 100 100 Tm " : "1 0 0 1 100 100 Tm ";
        string stream = clip + "BT /F1 10 Tf " + matrix + "2 Tc (ABCD) Tj ET";
        var reference = new StringBuilder(clip);
        for (int index = 0; index < 4; index++) reference.Append("BT /F1 10 Tf ").Append(matrix)
            .Append(index * 8).Append(" 0 Td (").Append((char)('A' + index)).Append(") Tj ET ");
        OfficeDrawing actual = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(stream, resources, font));
        OfficeDrawing expected = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(reference.ToString(), resources, font));
        byte[] pixels = OfficeDrawingRasterRenderer.Render(actual).GetPixels();
        Assert.Contains(pixels, channel => channel != 0);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected).GetPixels(), pixels);
    }

    [Theory]
    [InlineData("", 1000)]
    [InlineData("0 0 1 1 re W n 10 10 1 1 re W n ", 20)]
    [InlineData("0 0 1 1 re W n ", 20)]
    public void RenderPage_NonpaintingSpacedTextDoesNotConsumePositioningBudget(string clip, int x) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        string text = new string('A', PdfReadLimits.Default.MaxPositionedTextCharactersPerPage + 1);
        byte[] pdf = BuildSingleStreamPdf(clip + "BT /F1 10 Tf 1 Tc " + x + " 100 Td (" + text + ") Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = 1 }
        });
        OfficeDrawing drawing = document.Pages[0].ToDrawing();
        Assert.Empty(drawing.Elements);
    }

    [Theory]
    [InlineData(-70, 1)]
    [InlineData(300, -13)]
    public void RenderPage_PartiallyVisibleRunChargesOnlyProjectedGlyphs(int x, int spacing) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        const string resources = "<< /Font << /F1 5 0 R >> >>";
        string value = new string('A', 100);
        byte[] pdf = BuildSingleStreamPdf($"BT /F1 10 Tf {spacing} Tc {x} 100 Td ({value}) Tj ET", resources, font);
        const int visibleGlyphs = 35;
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = visibleGlyphs }
        });
        OfficeDrawing actual = document.Pages[0].ToDrawing();
        var reference = new StringBuilder();
        for (int index = 0; index < value.Length; index++) {
            int origin = x + index * (6 + spacing);
            reference.Append($"BT /F1 10 Tf {origin} 100 Td (A) Tj ET ");
        }
        OfficeDrawing expected = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(reference.ToString(), resources, font));
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected).GetPixels(), OfficeDrawingRasterRenderer.Render(actual).GetPixels());
        PdfReadDocument limited = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = visibleGlyphs - 1 }
        });
        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => limited.Pages[0].ToDrawing());
        Assert.Equal(PdfReadLimitKind.PositionedTextCharacters, error.Kind);
        Assert.Equal(visibleGlyphs, error.Actual);
    }

    [Fact]
    public void SpacedTextProjectionObservesInvocationCancellationWhenCallerOmitsToken() {
        using var cancellation = new System.Threading.CancellationTokenSource();
        PdfReadPage page = PdfReadDocument.Open(BuildSingleStreamPdf("", "<< >>")).Pages[0];
        var budget = new PdfReadPage.PageContentBudget(page, cancellation.Token);
        var span = new PdfTextSpan("AB", "F1", 10, 20, 100, 14, OfficeColor.Black, true, 0D, "Courier", null,
            characterAdvances: new[] { 7D, 7D }, canScaleAggregateAdvance: false,
            glyphCharacterLengths: new[] { 1, 1 }, glyphPaintedAdvances: new CancellingGlyphWidths(cancellation));
        var drawing = new OfficeDrawing(240, 200);
        var method = typeof(PdfReadPage).GetMethod("AddTextSpan", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;
        var exception = Assert.Throws<System.Reflection.TargetInvocationException>(() => method.Invoke(null,
            new object[] { drawing, 200D, span, budget, default(System.Threading.CancellationToken) }));
        Assert.IsType<OperationCanceledException>(exception.InnerException);
        Assert.Empty(drawing.Elements);
    }

    private sealed class CancellingGlyphWidths : System.Collections.Generic.IReadOnlyList<double> {
        private readonly System.Threading.CancellationTokenSource _cancellation;
        internal CancellingGlyphWidths(System.Threading.CancellationTokenSource cancellation) => _cancellation = cancellation;
        public int Count => 2;
        public double this[int index] { get { _cancellation.Cancel(); return 6D; } }
        public System.Collections.Generic.IEnumerator<double> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    [Theory]
    [InlineData("06280628", "بب")]
    [InlineData("0915093F", "कि")]
    [InlineData("00610301", "a\u0301")]
    public void RenderPage_SpacingRetainsContextualTextRuns(string unicode, string expectedText) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier /ToUnicode 6 0 R >>\nendobj";
        string cmap = "2 beginbfchar\n<41> <" + unicode.Substring(0, 4) + ">\n<42> <" + unicode.Substring(4) + ">\nendbfchar";
        byte[] pdf = BuildSingleStreamPdf("BT /F1 20 Tf 2 Tw 20 100 Td (AB) Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font, BuildStreamObject(6, "<<", cmap));
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = 1 }
        });
        PdfTextSpan span = Assert.Single(document.Pages[0].GetTextSpans());
        Assert.Equal(expectedText, span.Text);
        Assert.False(span.CanScaleAggregateAdvance);
        OfficeDrawing drawing = document.Pages[0].ToDrawing();
        // No word separator occurs: Tw must not change shaping, nor consume an expansion budget.
        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(expectedText, text.Text);
        var expected = new OfficeDrawing(drawing.Width, drawing.Height).AddText(text.Text, text.X, text.Y,
            text.Width, text.Height, text.Font, text.Color, wrapText: false);
        expected.Fonts.AddRange(drawing.Fonts);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected).GetPixels(), OfficeDrawingRasterRenderer.Render(drawing).GetPixels());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RenderPage_PositioningBudgetCountsPaintedExpansionOnly(bool actualText) {
        string font = actualText
            ? "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj"
            : "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier /FirstChar 65 /LastChar 66 /Widths [0 0] >>\nendobj";
        string content = actualText ? "/Span << /ActualText (Long replacement text) >> BDC " : "";
        content += "BT /F1 10 Tf 1 Tc 20 100 Td (AB) Tj ET";
        if (actualText) content += " EMC";
        PdfReadDocument document = PdfReadDocument.Open(BuildSingleStreamPdf(content, "<< /Font << /F1 5 0 R >> >>", font),
            new PdfLoadOptions { Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = actualText ? 2 : 1 } });
        // ActualText replaces extraction, while the display retains the two original glyphs.
        // Zero-width glyphs use the aggregate fallback and consume no expansion budget.
        Assert.Equal(actualText ? 2 : 1, document.Pages[0].ToDrawing().Elements.OfType<OfficeDrawingText>().Count());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RenderPage_SpacedGlyphBudgetIsSharedAcrossRunsAndForms(bool splitRuns, bool useForm) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        const string resources = "<< /Font << /F1 5 0 R >> >>";
        string content = "BT /F1 10 Tf 1 Tc 20 100 Td " +
            (splitRuns ? "(AB) Tj (CD) Tj" : "(ABCD) Tj") + " ET";
        byte[] pdf = useForm
            ? BuildSingleStreamPdf("/Fm Do q 1 0 0 1 1000 0 cm /Fm Do Q /Fm Do", "<< /XObject << /Fm 6 0 R >> >>", font,
                BuildStreamObject(6, "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources " + resources, content))
            : BuildSingleStreamPdf(content, resources, font);
        int exactCost = useForm ? 8 : 4;
        PdfReadDocument limited = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = exactCost - 1 }
        });
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => limited.Pages[0].ToDrawing());
        Assert.Equal(PdfReadLimitKind.PositionedTextCharacters, exception.Kind);
        Assert.Equal(exactCost - 1, exception.Limit);
        Assert.Equal(exactCost, exception.Actual);

        PdfReadDocument allowed = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = exactCost }
        });
        Assert.NotNull(allowed.Pages[0].ToDrawing());
        Assert.NotNull(allowed.Pages[0].ToDrawing()); // The budget belongs to one render invocation.
    }

    [Fact]
    public void RenderPage_DefaultBudgetRejectsOversizedSpacedRun() {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        string text = new string('A', PdfReadLimits.Default.MaxPositionedTextCharactersPerPage + 1);
        byte[] pdf = BuildSingleStreamPdf("BT /F1 10 Tf -6 Tc 20 100 Td (" + text + ") Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfPageImageRenderer.RenderPage(pdf));
        Assert.Equal(PdfReadLimitKind.PositionedTextCharacters, exception.Kind);
        Assert.Equal(text.Length, exception.Actual);
    }

    [Theory]
    [InlineData("87,0", -0.5, 0, false, false)]
    [InlineData("73%", 2, 0, false, false)]
    [InlineData("A B", 0, 4, false, false)]
    [InlineData("87,0", -0.5, 0, true, false)]
    [InlineData("A B", 2, 4, true, false)]
    [InlineData("87,0", -0.5, 0, false, true)]
    [InlineData("A B", 2, 4, true, true)]
    [InlineData("87,0", -8, 0, false, false)]
    [InlineData("87,0", -6, 0, false, false)]
    public void RenderPage_SpacedTextMatchesIndividuallyPositionedGlyphs(
        string value, double characterSpacing, double wordSpacing, bool clipped, bool rotated) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        const string resources = "<< /Font << /F1 5 0 R >> >>";
        string clip = clipped ? "90 80 26 30 re W n " : string.Empty;
        string matrix = rotated ? "0 1 -1 0 100 100 Tm " : "1 0 0 1 100 100 Tm ";
        string stream = clip + "BT /F1 10 Tf " + matrix +
            FormattableString.Invariant($"{characterSpacing} Tc {wordSpacing} Tw ({value}) Tj ET");

        // Courier has a 600-unit advance. Spacing moves the next glyph, while each
        // painted glyph remains six points wide. Explicit origins form the reference.
        var reference = new StringBuilder(clip);
        double offset = 0;
        foreach (char character in value) {
            reference.Append("BT /F1 10 Tf ").Append(matrix)
                .Append(offset.ToString("R", CultureInfo.InvariantCulture)).Append(" 0 Td (")
                .Append(character).Append(") Tj ET ");
            offset += 6 + characterSpacing + (character == ' ' ? wordSpacing : 0);
        }

        OfficeDrawing actual = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(stream, resources, font));
        OfficeDrawing expected = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(reference.ToString(), resources, font));
        byte[] expectedPixels = OfficeDrawingRasterRenderer.Render(expected).GetPixels();
        Assert.Contains(expectedPixels, channel => channel != 0);
        Assert.Equal(expectedPixels, OfficeDrawingRasterRenderer.Render(actual).GetPixels());
    }

    [Fact]
    public void ExportImage_ExcelSummaryKeepsEveryDigitAndPercentage() {
        // Desktop Excel PDF from the documented two-service operational dashboard example.
        string path = Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Rendering", "excel-operational-summary.pdf");
        PdfReadDocument document = PdfReadDocument.Open(File.ReadAllBytes(path));
        Assert.Contains("87,0", document.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("73%", document.Pages[0].ExtractText(), StringComparison.Ordinal);

        var exported = document.Pages[0].ExportImage(OfficeImageExportFormat.Png);
        Assert.True(OfficePngReader.TryDecode(exported.Bytes, out OfficeRasterImage? image));
        foreach (var bounds in new[] {
            (Left: 221, Top: 86, Right: 225, Bottom: 95),
            (Left: 225, Top: 86, Right: 229, Bottom: 95),
            (Left: 229, Top: 86, Right: 231, Bottom: 95),
            (Left: 231, Top: 86, Right: 235, Bottom: 95),
            (Left: 221, Top: 107, Right: 225, Bottom: 115),
            (Left: 225, Top: 107, Right: 229, Bottom: 115),
            (Left: 229, Top: 107, Right: 236, Bottom: 115)
        }) {
            int painted = 0;
            for (int y = bounds.Top; y < bounds.Bottom; y++) {
                for (int x = bounds.Left; x < bounds.Right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A > 0 && pixel.R < 180 && pixel.G < 180 && pixel.B < 180) painted++;
                }
            }
            Assert.True(painted > 0, $"Missing glyph in {bounds}.");
        }
    }
}
