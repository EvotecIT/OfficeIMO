using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingTextOpacityTests {
    [Theory]
    [InlineData("rich", (byte)0)]
    [InlineData("rich", (byte)128)]
    [InlineData("positioned", (byte)0)]
    [InlineData("positioned", (byte)128)]
    [InlineData("rotated", (byte)0)]
    [InlineData("rotated", (byte)128)]
    [InlineData("clipped", (byte)0)]
    [InlineData("clipped", (byte)128)]
    public void DrawingForegroundAlphaPreservesLogicalTextAndLayout(string path, byte alpha) {
        byte[] bytes = CreatePdf(CreateDrawing(path, alpha));
        byte[] opaqueBytes = CreatePdf(CreateDrawing(path, 255));
        string raw = Encoding.ASCII.GetString(bytes);
        string opacity = alpha == 0 ? "0" : "0.502";

        Assert.Contains("/ca " + opacity + " /CA 1", raw, StringComparison.Ordinal);
        Assert.Contains(" gs", raw, StringComparison.Ordinal);
        PdfTextSpan[] spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan[] opaqueSpans = PdfReadDocument.Open(opaqueBytes).Pages[0].GetTextSpans().ToArray();
        Assert.Equal(opaqueSpans.Length, spans.Length);
        Assert.Contains("Alpha", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        for (int index = 0; index < spans.Length; index++) {
            Assert.Equal(opaqueSpans[index].Text, spans[index].Text);
            Assert.Equal(opaqueSpans[index].X, spans[index].X, 6);
            Assert.Equal(opaqueSpans[index].Y, spans[index].Y, 6);
            Assert.Equal(opaqueSpans[index].Advance, spans[index].Advance, 6);
            Assert.Equal(opaqueSpans[index].RotationDegrees, spans[index].RotationDegrees, 6);
            Assert.Equal(alpha, spans[index].Color!.Value.A);
        }

        // An independent text consumer still sees the same glyph positions and source text.
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        using var opaquePdf = UglyToad.PdfPig.PdfDocument.Open(opaqueBytes);
        var letters = pdf.GetPage(1).Letters;
        var opaqueLetters = opaquePdf.GetPage(1).Letters;
        Assert.Equal("Alpha", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(opaqueLetters.Count, letters.Count);
        for (int index = 0; index < letters.Count; index++) {
            Assert.Equal(opaqueLetters[index].StartBaseLine.X, letters[index].StartBaseLine.X, 6);
            Assert.Equal(opaqueLetters[index].StartBaseLine.Y, letters[index].StartBaseLine.Y, 6);
            Assert.Equal(opaqueLetters[index].Width, letters[index].Width, 6);
        }
    }

    [Theory]
    [InlineData((byte)0)]
    [InlineData((byte)128)]
    public void RichRunAlphaReachesDefaultDecorationsWithoutFadingBackgroundOrFollowingRun(byte alpha) {
        var drawing = new OfficeDrawing(240D, 160D).AddRichText(new[] {
            new OfficeRichTextRun("Alpha", 20D, OfficeColor.FromRgba(255, 0, 0, alpha),
                underline: true, strikethrough: true, fontFamily: "Helvetica", backgroundColor: OfficeColor.Yellow),
            new OfficeRichTextRun("Next", 20D, OfficeColor.Blue, fontFamily: "Helvetica")
        }, 20D, 20D, 200D, 60D);
        byte[] bytes = CreatePdf(drawing);
        var spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans();

        Assert.Equal(alpha, Assert.Single(spans, span => span.Text == "Alpha").Color!.Value.A);
        Assert.Equal(OfficeColor.Blue, Assert.Single(spans, span => span.Text == "Next").Color);
        OfficeDrawing rendered = PdfPageImageRenderer.RenderPage(bytes);
        var shapes = EnumerateElements(rendered).OfType<OfficeDrawingShape>().Select(item => item.Shape).ToArray();
        Assert.Contains(shapes, shape => shape.FillColor == OfficeColor.Yellow && (shape.FillOpacity ?? 1D) == 1D);
        if (alpha == 0) {
            Assert.DoesNotContain(shapes, shape => shape.StrokeColor.HasValue && (shape.StrokeOpacity ?? 1D) > 0D);
        } else {
            Assert.Equal(2, shapes.Count(shape => shape.StrokeColor == OfficeColor.Red &&
                Math.Abs((shape.StrokeOpacity ?? 1D) - alpha / 255D) < 0.001D));
        }
        Assert.Contains("/ca 1 /CA " + (alpha == 0 ? "0" : "0.502"), Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitDecorationAlphaIsIndependentOfTransparentForeground(bool wrapText) {
        var drawing = new OfficeDrawing(240D, 160D).AddStyledText(
            "Alpha", 20D, 20D, 180D, 60D, new OfficeFontInfo("Helvetica", 20D),
            OfficeColor.FromRgba(255, 0, 0, 0), OfficeTextAlignment.Left, null, OfficeTextVerticalAlignment.Top,
            0D, null, null, wrapText, false, false, false, false, null, null,
            OfficeTextDecorationStyle.Single, OfficeTextDecorationStyle.Single, OfficeTextBaseline.Normal,
            baselineLevel: 0, baselineScale: 1D, baselineOffset: 0D, decorationColor: OfficeColor.FromRgba(0, 0, 255, 128));
        byte[] bytes = CreatePdf(drawing);
        var shapes = EnumerateElements(PdfPageImageRenderer.RenderPage(bytes))
            .OfType<OfficeDrawingShape>().Select(item => item.Shape).ToArray();

        Assert.Equal(2, shapes.Count(shape => shape.StrokeColor == OfficeColor.Blue &&
            Math.Abs((shape.StrokeOpacity ?? 1D) - 128D / 255D) < 0.001D));
        Assert.Contains("Alpha", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/ca 0 /CA 1", raw, StringComparison.Ordinal);
        Assert.Contains("/ca 1 /CA 0.502", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void DrawingForegroundAlphaRetainsIsolatedGroupOpacityAndRestoresOutsideGroup() {
        OfficeDrawing child = CreateDrawing("rich", 128);
        var drawing = new OfficeDrawing(260D, 180D)
            .AddEffectDrawing(child, OfficeTransform.Translate(10D, 5D), opacity: 0.5D)
            .AddPositionedText("Next", 20D, 110D, 80D, 30D, new OfficeFontInfo("Helvetica", 20D), OfficeColor.Blue);
        byte[] bytes = CreatePdf(drawing);
        var spans = PdfReadDocument.Open(bytes).Pages[0].GetTextSpans();

        // Text spans expose local glyph paint; the PDF transparency group
        // composites that paint separately through its own graphics state.
        Assert.Equal((byte)128, Assert.Single(spans, span => span.Text == "Alpha").Color!.Value.A);
        Assert.Equal(OfficeColor.Blue, Assert.Single(spans, span => span.Text == "Next").Color);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Group << /S /Transparency /I true /K false >>", raw, StringComparison.Ordinal);
        Assert.Contains("/ca 0.5 /CA 0.5", raw, StringComparison.Ordinal);
        Assert.Contains("/ca 0.502 /CA 1", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void StackedVerticalFallbackRetainsForegroundAlphaAndLogicalText() {
        var drawing = new OfficeDrawing(240D, 160D)
            .AddVerticalText("Alpha", 20D, 10D, 80D, 140D,
                new OfficeFontInfo("Helvetica", 20D), OfficeColor.FromRgba(255, 0, 0, 128));
        byte[] bytes = CreatePdf(drawing);

        Assert.Contains("Alpha", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Contains("/ca 0.502 /CA 1", Encoding.ASCII.GetString(bytes), StringComparison.Ordinal);
    }

    private static OfficeDrawing CreateDrawing(string path, byte alpha) {
        var color = OfficeColor.FromRgba(255, 0, 0, alpha);
        var drawing = new OfficeDrawing(240D, 160D);
        switch (path) {
            case "rich":
                return drawing.AddRichText(new[] { new OfficeRichTextRun("Alpha", 20D, color, fontFamily: "Helvetica") },
                    20D, 20D, 180D, 60D);
            case "positioned":
                return drawing.AddPositionedText("Alpha", 20D, 20D, 180D, 60D,
                    new OfficeFontInfo("Helvetica", 20D), color, textAdvanceWidth: 80D);
            case "rotated":
                return drawing.AddPositionedText("Alpha", 20D, 20D, 180D, 60D,
                    new OfficeImageFrameTransform(15D, 110D, 50D),
                    new OfficeFontInfo("Helvetica", 20D), color, textAdvanceWidth: 80D);
            case "clipped":
                return drawing.AddClippedDrawing(CreateDrawing("rich", alpha), 0D, 0D,
                    OfficeClipPath.Rectangle(220D, 100D));
            default:
                throw new ArgumentOutOfRangeException(nameof(path));
        }
    }

    private static byte[] CreatePdf(OfficeDrawing drawing) => PdfDocument.Create(new PdfOptions {
        PageWidth = drawing.Width, PageHeight = drawing.Height,
        MarginLeft = 0D, MarginTop = 0D, MarginRight = 0D, MarginBottom = 0D,
        CompressContentStreams = false
    }).Compose(composer => composer.Page(page => page.Content(content => content.Drawing(drawing)))).ToBytes();

    private static IEnumerable<OfficeDrawingElement> EnumerateElements(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            yield return element;
            OfficeDrawing? child = element switch {
                OfficeDrawingGroup group => group.Drawing,
                OfficeDrawingEffectGroup effect => effect.Drawing,
                _ => null
            };
            if (child != null) {
                foreach (OfficeDrawingElement nested in EnumerateElements(child)) yield return nested;
            }
        }
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(128, false)]
    [InlineData(128, true)]
    public void DrawingTextRetainsAlphaWithoutChangingTheFollowingPaint(byte alpha, bool wrap) {
        var drawing = new OfficeDrawing(180D, 110D);
        drawing.AddText("FIRST", 0D, 0D, 180D, 40D, new OfficeFontInfo("Arial", 24D),
            OfficeColor.FromRgba(255, 0, 0, alpha), wrapText: wrap);
        drawing.AddText("FOLLOW", 0D, 60D, 180D, 40D, new OfficeFontInfo("Arial", 24D), OfficeColor.FromRgb(0, 0, 255));
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 220D, PageHeight = 150D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
        });
        document.Content.Canvas(canvas => canvas.Drawing(drawing, 0D, 0D, 180D, 110D));
        byte[] bytes = document.ToBytes();
        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("FIRST", text);
        Assert.Contains("FOLLOW", text);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes), 1D, OfficeColor.White);
        int firstInk = 0;
        int followingBlue = 0;
        for (int y = 20; y < 60; y++) {
            for (int x = 20; x < 200; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.R == 255 && pixel.G == 255 && pixel.B == 255) continue;
                firstInk++;
                Assert.Equal(255, pixel.R);
                Assert.True(pixel.G >= 126 && pixel.B >= 126, "Half-transparent red text must blend with the white page.");
            }
        }
        for (int y = 80; y < 120; y++)
            for (int x = 20; x < 200; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.B > 240 && pixel.R < 30 && pixel.G < 30) followingBlue++;
            }
        Assert.Equal(alpha == 0, firstInk == 0);
        Assert.True(followingBlue > 0, "The following opaque drawing text must retain its own paint state.");
    }
    [Theory]
    [InlineData(false, 128, 255)]
    [InlineData(true, 128, 255)]
    [InlineData(false, 255, 128)]
    [InlineData(true, 255, 128)]
    public void IndependentDecorationAlphaIsPreserved(bool wrap, byte textAlpha, byte decorationAlpha) {
        var drawing = new OfficeDrawing(180D, 110D);
        drawing.AddStyledText("FIRST", 0D, 0D, 180D, 40D, new OfficeFontInfo("Arial", 24D),
            OfficeColor.FromRgba(255, 0, 0, textAlpha), OfficeTextAlignment.Left, null,
            OfficeTextVerticalAlignment.Top, 0D, null, null, wrap, false, false, false, false,
            null, null, OfficeTextDecorationStyle.Single, OfficeTextDecorationStyle.Single,
            OfficeTextBaseline.Normal, 0, 1D, 0D, OfficeColor.FromRgba(0, 0, 255, decorationAlpha));
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 220D, PageHeight = 150D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
        });
        document.Content.Canvas(canvas => canvas.Drawing(drawing, 0D, 0D, 180D, 110D));
        OfficeDrawing scene = PdfPageImageRenderer.RenderPage(document.ToBytes());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(scene, 2D, OfficeColor.White);
        int opaqueBlue = 0;
        int translucentBlue = 0;
        for (int y = 40; y < 120; y++)
            for (int x = 40; x < 400; x++) {
                OfficeColor pixel = raster.GetPixel(x, y);
                if (pixel.B > 240 && pixel.R < 100 && pixel.G < 100) opaqueBlue++;
                if (pixel.B > 240 && pixel.R >= 100 && pixel.R < 230 && pixel.G >= 100 && pixel.G < 230) translucentBlue++;
            }
        Assert.Equal(decorationAlpha == 255, opaqueBlue > 0);
        if (decorationAlpha == 128) Assert.True(translucentBlue > 0, "Independent decoration paint must remain visible.");
    }
    [Theory]
    [InlineData(false, 128, 255)]
    [InlineData(true, 128, 255)]
    [InlineData(false, 255, 128)]
    [InlineData(true, 255, 128)]
    public void TaggedTransformedTextAppliesAlphaInsideIsolatedPaint(bool wrap, byte glyphAlpha, byte strokeAlpha) {
        var drawing = new OfficeDrawing(180D, 110D);
        drawing.AddStyledText("FIRST", 0D, 0D, 180D, 40D, new OfficeFontInfo("Arial", 24D),
            OfficeColor.FromRgba(255, 0, 0, glyphAlpha), OfficeTextAlignment.Left, null,
            OfficeTextVerticalAlignment.Top, 5D, null, null, wrap, false, false, false, false,
            null, null, OfficeTextDecorationStyle.Single, OfficeTextDecorationStyle.Single,
            OfficeTextBaseline.Normal, 0, 1D, 0D, OfficeColor.FromRgba(0, 0, 255, strokeAlpha));
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 220D, PageHeight = 150D, MarginLeft = 20D, MarginRight = 20D,
            MarginTop = 20D, MarginBottom = 20D, CompressContentStreams = false,
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
        });
        document.Content.Canvas(canvas => canvas.Drawing(drawing, 0D, 0D, 180D, 110D));
        string raw = System.Text.Encoding.GetEncoding(28591).GetString(document.ToBytes());
        var forms = System.Text.RegularExpressions.Regex.Matches(raw,
            @"(?s)/Subtype /Form(?<dictionary>.*?)stream\r?\n(?<content>.*?)\r?\nendstream");
        var form = Assert.Single(forms.Cast<System.Text.RegularExpressions.Match>());
        Assert.Contains("/I true", form.Groups["dictionary"].Value);
        // Isolated transparency Forms reset alpha; a caller gs cannot express separate
        // glyph/decorative alpha. Verify the state used by the inner paint, not our reader.
        var states = System.Text.RegularExpressions.Regex.Matches(form.Groups["content"].Value, @"/(?<name>GS\d+) gs");
        Assert.NotEmpty(states.Cast<System.Text.RegularExpressions.Match>());
        string name = states[0].Groups["name"].Value;
        var reference = System.Text.RegularExpressions.Regex.Match(form.Groups["dictionary"].Value,
            "/" + name + @" (?<id>\d+) 0 R");
        Assert.True(reference.Success);
        var state = System.Text.RegularExpressions.Regex.Match(raw,
            reference.Groups["id"].Value + @" 0 obj\s*<< /Type /ExtGState /ca (?<fill>[\d.]+) /CA (?<stroke>[\d.]+)");
        Assert.True(state.Success);
        Assert.Equal(glyphAlpha / 255D, double.Parse(state.Groups["fill"].Value, System.Globalization.CultureInfo.InvariantCulture), 3);
        Assert.Equal(strokeAlpha / 255D, double.Parse(state.Groups["stroke"].Value, System.Globalization.CultureInfo.InvariantCulture), 3);
    }

}
