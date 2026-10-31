using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSyntheticFontWeightTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    public void MissingWeightPaintsBoldWithoutChangingAdjacentText(bool named, bool italic, bool physicalBold) {
        byte[] regular = Font("Regular");
        var family = new PdfEmbeddedFontFamily("Weight proof", regular,
            bold: physicalBold ? Font("Bold") : null, italic: italic ? Font("Italic") : null);
        var options = new PdfCreateOptions { EmbedDefaultFonts = false, DefaultTextColor = OfficeColor.Black };
        if (named) options.RegisterNamedFontFamily(family);
        else options.UseFontFamily(family);
        OfficeColor ink = OfficeColor.FromRgb(25, 140, 77);
        using PdfDocument source = PdfDocument.Create(options).Compose(composition => composition.Content(content => content
            .Paragraph(paragraph => {
                if (named) paragraph.FontFamily("Weight proof");
                paragraph.FontSize(OfficeLength.Points(24)).Text("LEFT ").Bold(true).Italic(italic).Color(ink).Text("BOLD")
                    .Bold(false).Italic(false).ResetColor().Text(" RIGHT");
            })
            .Paragraph(paragraph => paragraph.Text("AFTER"))));
        byte[] bytes = source.ToBytes();
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(bytes).Pages);
        PdfTextSpan heavy = Assert.Single(page.GetTextSpans(), span => span.Text.Contains("BOLD", StringComparison.Ordinal));
        Assert.Equal(physicalBold ? 0 : 2, heavy.TextRenderingMode);
        Assert.Equal(ink, heavy.Color);
        if (!physicalBold) {
            Assert.True(heavy.OutlineStroke.HasValue);
            Assert.Equal(ink, heavy.OutlineStroke.Value.Color);
            Assert.Equal(1D, heavy.OutlineStroke.Value.Width, 3);
        }
        Assert.All(page.GetTextSpans().Where(span => !span.Text.Contains("BOLD", StringComparison.Ordinal)), span => {
            Assert.Equal(0, span.TextRenderingMode);
            Assert.Equal(OfficeColor.Black, span.Color);
        });
        using var independent = PdfPigDocument.Open(bytes);
        string[] words = independent.GetPage(1).GetWords().Select(word => word.Text).ToArray();
        Assert.Equal(new[] { "LEFT", "BOLD", "RIGHT", "AFTER" }, words);
        Assert.NotEmpty(OfficeDrawingRasterRenderer.ToPng(page.ToDrawing()));
    }

    [Theory]
    [InlineData((byte)0, false)]
    [InlineData((byte)128, false)]
    [InlineData((byte)128, true)]
    public void SyntheticDrawingWeightUsesTheForegroundAlpha(byte alpha, bool tagged) {
        var options = new PdfCreateOptions { EmbedDefaultFonts = false, PageWidth = OfficeLength.Points(300), PageHeight = OfficeLength.Points(120),
            MarginLeft = OfficeLength.Points(0), MarginTop = OfficeLength.Points(0), MarginRight = OfficeLength.Points(0), MarginBottom = OfficeLength.Points(0) }.RegisterNamedFontFamily(
            new PdfEmbeddedFontFamily("Weight proof", Font("Regular")));
        if (tagged) options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
        var drawing = new OfficeDrawing(240, 100).AddRichText(new[] {
            new OfficeRichTextRun("BOLD", 24D, OfficeColor.FromRgba(25, 140, 77, alpha),
                bold: true, fontFamily: "Weight proof")
        }, 20, 20, 200, 60);
        using PdfDocument document = PdfDocument.Create(options).Compose(composition => composition.Content(content =>
            content.Drawing(drawing)));
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(document.ToBytes()).Pages[0].GetTextSpans(),
            item => item.Text == "BOLD");
        Assert.Equal(alpha, span.Color!.Value.A);
        Assert.Equal(alpha, span.OutlineStroke?.Color.A ?? 0);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(document.ToBytes()), 1D, OfficeColor.White);
        bool painted = false;
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++)
                if (image.GetPixel(x, y) != OfficeColor.White) {
                    painted = true;
                    // One translucent glyph has one opacity, including its added
                    // weight. Repainting its edge must not create a darker ring.
                    Assert.True(image.GetPixel(x, y).G >= 195, image.GetPixel(x, y).ToString());
                }
        Assert.Equal(alpha != 0, painted);
    }

    [Fact]
    public void WeightAddedMidParagraphRestoresTheSurroundingStrokeState() {
        var stream = new System.Text.StringBuilder();
        new ContentStreamBuilder(stream).LineWidth(7).StrokeColor(OfficeColor.FromRgb(255, 0, 0)).StrokeDash(4D, 2D)
            .BeginText().Font("F1", 24).TextMatrix(20, 50).ShowHexText("41")
            .Font("F1", 24, syntheticBold: true).TextMatrix(50, 50).FillColor(OfficeColor.FromRgb(0, 0, 255)).ShowHexText("42")
            .Font("F1", 24).TextMatrix(80, 50).ShowHexText("43").EndText()
            .MoveTo(10, 20).LineTo(100, 20).StrokePath();
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(NativePage(stream.ToString())).Pages);
        PdfTextSpan[] text = page.GetTextSpans().ToArray();
        Assert.Equal(new[] { "A", "B", "C" }, text.Select(span => span.Text));
        Assert.Equal(new[] { 0, 2, 0 }, text.Select(span => span.TextRenderingMode));
        Assert.Equal(new[] { 20D, 50D, 80D }, text.Select(span => span.X));
        Assert.True(text[1].OutlineStroke.HasValue);
        Assert.Empty(text[1].OutlineStroke!.Value.DashPattern ?? Array.Empty<double>());
        OfficeShape line = Assert.Single(page.ToDrawing().Shapes, item => item.Shape.Kind == OfficeShapeKind.Line).Shape;
        Assert.Equal(OfficeColor.Red, line.StrokeColor);
        Assert.Equal(7D, line.StrokeWidth);
        Assert.Equal(new[] { 4D, 2D }, line.StrokeDashArray);
    }

    [Theory]
    [InlineData(0.12D, false)]
    [InlineData(0.5D, false)]
    [InlineData(0.5D, true)]
    public void SyntheticWatermarkAppliesOpacityOnce(double opacity, bool tagged) {
        var options = new PdfCreateOptions {
            EmbedDefaultFonts = false, PageWidth = OfficeLength.Points(300), PageHeight = OfficeLength.Points(120),
            MarginLeft = OfficeLength.Points(0), MarginTop = OfficeLength.Points(0), MarginRight = OfficeLength.Points(0), MarginBottom = OfficeLength.Points(0),
            TextWatermark = new PdfTextWatermark("WAVE") {
                FontSize = OfficeLength.Points(54D), RotationAngle = 0D, Color = OfficeColor.FromRgb(25, 140, 77), Opacity = opacity
            }
        }.UseFontFamily(new PdfEmbeddedFontFamily("Weight proof", Font("Regular")));
        if (tagged) options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers;
        using PdfDocument document = PdfDocument.Create(options).Compose(c => c.Content(t => t.Paragraph(p => p.Color(OfficeColor.White).Text("BODY"))));
        byte[] bytes = document.ToBytes();
        Assert.Contains("WAVE", PdfReadDocument.Open(bytes, new PdfLoadOptions { IncludeArtifactText = true }).ExtractText(), StringComparison.Ordinal);
        if (tagged) Assert.DoesNotContain("WAVE", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        using var independent = PdfPigDocument.Open(bytes);
        Assert.Contains("WAVE", independent.GetPage(1).GetWords().Select(word => word.Text));
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(bytes), 1D, OfficeColor.White);
        int minimumGreen = (int)Math.Floor(255D * (1D - opacity) + 140D * opacity) - 2;
        bool painted = false;
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel == OfficeColor.White) continue;
                painted = true;
                Assert.True(pixel.G >= minimumGreen, pixel.ToString());
            }
        Assert.True(painted);
    }

    [Theory]
    [InlineData(0, OfficeStrokeLineJoin.Miter, "miter")]
    [InlineData(1, OfficeStrokeLineJoin.Round, "round")]
    [InlineData(2, OfficeStrokeLineJoin.Bevel, "bevel")]
    public void TextOutlineRetainsNativeStrokeGeometry(int nativeJoin, OfficeStrokeLineJoin join, string svgJoin) {
        var content = new System.Text.StringBuilder();
        new ContentStreamBuilder(content).LineWidth(4D).LineCap(2).LineJoin(nativeJoin).MiterLimit(2D)
            .BeginText().Font("F1", 60D).TextMatrix(35D, 25D)
            .TextRenderingMode(2).ShowHexText("4156").EndText();
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(NativePage(content.ToString())).Pages);
        PdfTextSpan text = Assert.Single(page.GetTextSpans());
        Assert.Equal("AV", text.Text);
        Assert.Equal(OfficeStrokeLineCap.Square, text.OutlineStroke!.Value.Cap);
        Assert.Equal(join, text.OutlineStroke.Value.Join);
        Assert.Equal(2D, text.OutlineStroke.Value.MiterLimit);
        OfficeDrawing drawing = page.ToDrawing();
        string svg = System.Text.Encoding.UTF8.GetString(OfficeDrawingSvgExporter.ToSvgBytes(drawing, 1D, OfficeSvgSizeUnit.Pixel));
        Assert.Contains("stroke-linejoin=\"" + svgJoin + "\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-linecap=\"square\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-miterlimit=\"2\"", svg, StringComparison.Ordinal);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(drawing, 4D, OfficeColor.White);
        Assert.True(CountInk(actual) > 0);
        if (join == OfficeStrokeLineJoin.Round) {
            string miterContent = content.ToString().Replace("1 j", "0 j")
                .Replace("2 M", "10 M");
            OfficeRasterImage miter = OfficeDrawingRasterRenderer.Render(
                PdfReadDocument.Open(NativePage(miterContent)).Pages[0].ToDrawing(), 4D, OfficeColor.White);
            // Acute corners in A and V must keep their rounded outline instead
            // of growing the longer miter spikes of the alternate source state.
            Assert.True(CountInk(miter) > CountInk(actual));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeFormTextInheritsTheRestoredStrokeState(bool graphicsState) {
        string parent = graphicsState ? "/GS gs\n" : "4 w 2 J 1 j 2 M [3 3] 0 d\n";
        // A local override must not leak past Q, including into nested Forms.
        parent += "q 0 J 0 j 10 M Q /Fm Do\n";
        byte[] bytes = NativePage(parent, "BT /F1 60 Tf 2 Tr 1 0 0 1 35 25 Tm <4156> Tj ET\n");
        PdfReadPage page = Assert.Single(PdfReadDocument.Open(bytes).Pages);
        PdfTextSpan text = Assert.Single(page.GetTextSpans());
        Assert.Equal("AV", text.Text);
        Assert.Equal(OfficeStrokeLineCap.Square, text.OutlineStroke!.Value.Cap);
        Assert.Equal(OfficeStrokeLineJoin.Round, text.OutlineStroke.Value.Join);
        Assert.Equal(2D, text.OutlineStroke.Value.MiterLimit);
        Assert.Equal(4D, text.OutlineStroke.Value.Width);
        Assert.Equal(new[] { 3D, 3D }, text.OutlineStroke.Value.DashPattern);
        string svg = System.Text.Encoding.UTF8.GetString(OfficeDrawingSvgExporter.ToSvgBytes(page.ToDrawing()));
        Assert.Contains("stroke-linecap=\"square\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-linejoin=\"round\"", svg, StringComparison.Ordinal);
        Assert.True(CountInk(OfficeDrawingRasterRenderer.Render(page.ToDrawing(), 2D, OfficeColor.White)) > 0);
    }

    private static int CountInk(OfficeRasterImage image) {
        int count = 0;
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++)
                if (image.GetPixel(x, y) != OfficeColor.White) count++;
        return count;
    }

    private static byte[] NativePage(string content, string? formText = null) {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >>" +
                (formText == null ? "" : " /XObject << /Fm 6 0 R >> /ExtGState << /GS 8 0 R >>") + " >> >>",
            "<< /Length " + System.Text.Encoding.ASCII.GetByteCount(content) + ">>\nstream\n" + content + "endstream",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
        };
        if (formText != null) {
            const string outer = "q /Inner Do Q\n";
            objects.Add("<< /Type /XObject /Subtype /Form /BBox [0 0 200 100] /Resources << /XObject << /Inner 7 0 R >> >> /Length " + outer.Length + ">>\nstream\n" + outer + "endstream");
            objects.Add("<< /Type /XObject /Subtype /Form /BBox [0 0 200 100] /Resources << /Font << /F1 5 0 R >> >> /Length " + formText.Length + ">>\nstream\n" + formText + "endstream");
            objects.Add("<< /Type /ExtGState /LW 4 /LC 2 /LJ 1 /ML 2 /D [[3 3] 0] >>");
        }
        var output = new System.Text.StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Count; index++) {
            offsets.Add(output.Length);
            output.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = output.Length;
        output.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) output.Append(offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        output.Append("trailer\n<< /Root 1 0 R /Size ").Append(objects.Count + 1).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return System.Text.Encoding.ASCII.GetBytes(output.ToString());
    }

    private static byte[] Font(string style) => File.ReadAllBytes(Path.Combine(
        AppContext.BaseDirectory, "Typography", "Carlito-" + style + ".ttf"));
}
