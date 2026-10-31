using System;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("shear", 1D, 0D, 0.5D, 1D, false)]
    [InlineData("nonuniform-shear", 0.6D, 0D, 0.3D, 1.2D, false)]
    [InlineData("rotated-shear", 0.8D, 0.6D, -0.2D, 1.1D, false)]
    [InlineData("destination-clip", 1D, 0D, 0.5D, 1D, true)]
    public void RenderPage_PreservesAffineTextOutlineAndDestinationClip(
        string scenario, double a, double b, double c, double d, bool clipped) {
        byte[] pdf = BuildAffineOutlineTextPdf(a, b, c, d, "B", clipped: clipped, logicalB: "00660069");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        OfficeDrawing projected = page.ToDrawing();
        var retained = Assert.Single(ReplayText(projected));
        Assert.Equal("fi", page.ExtractText().Trim());
        Assert.Equal("fi", retained.Text.Text);
        Assert.NotEqual("fi", retained.Text.RasterText);
        // A clipped child can inherit its font collection from the page scene.
        bool resolved = retained.Owner.Fonts.TryResolveFaceForText(retained.Text.RasterText,
            retained.Text.Font.FamilyName, retained.Text.Font.Style, out OfficeFontFace? face) ||
            projected.Fonts.TryResolveFaceForText(retained.Text.RasterText,
                retained.Text.Font.FamilyName, retained.Text.Font.Style, out face);
        Assert.True(resolved);
        Assert.True(face!.Program.TryGetGlyphMetrics(retained.Text.RasterText[0], out int glyph, out _));
        Assert.Equal(2, glyph);

        var expected = new OfficeDrawing(240D, 200D);
        if (clipped) {
            // Intersect the authored parallelogram with the destination-space
            // PDF clip [66,100]-[74,112], independently of reader transforms.
            AddAffineOutlinePolygon(expected, new[] {
                new OfficePoint(66D, 100D), new OfficePoint(72D, 100D), new OfficePoint(74D, 104D),
                new OfficePoint(74D, 112D), new OfficePoint(66D, 112D)
            });
        } else {
            AddAffineOutlinePolygon(expected, AffineOutlineCorners(a, b, c, d, 0D, 12D));
        }
        AssertAffineOutlineRaster(scenario, pdf, page, expected);
        if (scenario == "shear") AssertAffineOutlineSvgRaster(scenario, page, expected);
    }

    [Fact]
    public void RenderPage_AffineCharacterAndWordSpacingPreserveGlyphOriginsAndPositioningBudget() {
        const double a = 0.8D, b = 0.6D, c = -0.2D, d = 1.1D;
        byte[] pdf = BuildAffineOutlineTextPdf(a, b, c, d, "A B", characterSpacing: 2D, wordSpacing: 3D);
        PdfReadPage page = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = 2 }
        }).Pages[0];
        OfficeDrawing projected = page.ToDrawing();
        Assert.Equal(new[] { "A", "B" }, ReplayText(projected).Select(item => item.Text.Text));
        var expected = new OfficeDrawing(240D, 200D);
        AddAffineOutlinePolygon(expected, AffineOutlineCorners(a, b, c, d, 0D, 8D));
        // A advances 10+2 points; the declared space advances 5+2+3.
        // Tc/Tw move B's origin to 22 without stretching either outline.
        AddAffineOutlinePolygon(expected, AffineOutlineCorners(a, b, c, d, 22D, 12D));
        AssertAffineOutlineRaster("character-word-spacing", pdf, page, expected);

        PdfReadLimitException limit = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(pdf,
            new PdfLoadOptions { Limits = new PdfReadLimits { MaxPositionedTextCharactersPerPage = 1 } })
            .Pages[0].ToDrawing());
        Assert.Equal(PdfReadLimitKind.PositionedTextCharacters, limit.Kind);
        Assert.Equal(2, limit.Actual);
    }

    [Fact]
    public void RenderPage_AffineTextRetainsStrokeAndLogicalGlyphMetadataThroughCloning() {
        byte[] pdf = BuildAffineOutlineTextPdf(1D, 0D, 0.5D, 1D, "B", logicalB: "00660069", strokeOnly: true);
        OfficeDrawing drawing = PdfReadDocument.Open(pdf).Pages[0].ToDrawing();
        foreach (OfficeDrawing scene in new[] { drawing, drawing.Clone() }) {
            OfficeDrawingText text = Assert.Single(ReplayText(scene)).Text;
            Assert.Equal("fi", text.Text);
            Assert.NotEqual("fi", text.RasterText);
            Assert.Equal(OfficeColor.Blue, text.OutlineStrokeColor);
            Assert.Equal(0.75D, text.OutlineStrokeWidth, 6);
            Assert.Equal(new[] { 2D, 1D }, text.OutlineStrokeDashPattern);
            Assert.Equal(0.5D, text.OutlineStrokeDashOffset, 6);
            Assert.True(text.FillSuppressed);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RenderPage_AffineGlyphStrokeUsesDestinationWidthAndDash(bool dashed) {
        byte[] pdf = BuildAffineOutlineTextPdf(1D, 0D, 0.5D, 1D, "B", strokeOnly: true,
            strokeWidth: 8D, strokeDash: dashed ? "[9 5] 2" : "[] 0");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        OfficePoint[] points = AffineOutlineCorners(1D, 0D, 0.5D, 1D, 0D, 12D)
            .Select(point => new OfficePoint(point.X, 200D - point.Y)).ToArray();
        // Text-matrix geometry changes the outline before the PDF pen is
        // applied. The eight-point pen and its dash lengths stay in page units.
        var shape = OfficeShape.Polygon(points);
        shape.FillColor = null;
        shape.StrokeColor = OfficeColor.Blue;
        shape.StrokeWidth = 8D;
        shape.StrokeLineCap = OfficeStrokeLineCap.Butt;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Miter;
        shape.StrokeMiterLimit = 10D;
        if (dashed) shape.SetStrokeDashArray(new[] { 9D, 5D }, 2D);
        var expected = new OfficeDrawing(240D, 200D).AddShape(shape,
            points.Min(point => point.X), points.Min(point => point.Y));
        string scenario = dashed ? "stroke-dashed" : "stroke-solid";
        AssertAffineOutlineRaster(scenario, pdf, page, expected);
        AssertAffineOutlineSvgRaster(scenario, page, expected);
    }

    [Fact]
    public void ReauthoringAffineGlyphDrawingFailsExplicitlyWhileOrdinaryTextRemainsSupported() {
        OfficeDrawing Read(double shear) => PdfReadDocument.Open(
            BuildAffineOutlineTextPdf(1D, 0D, shear, 1D, "B")).Pages[0].ToDrawing();
        byte[] Write(OfficeDrawing drawing) => PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(240D), PageHeight = global::OfficeIMO.OfficeLength.Points(200D), MarginLeft = global::OfficeIMO.OfficeLength.Points(0D), MarginRight = global::OfficeIMO.OfficeLength.Points(0D), MarginTop = global::OfficeIMO.OfficeLength.Points(0D), MarginBottom = global::OfficeIMO.OfficeLength.Points(0D)
        }).Canvas(canvas => canvas.Drawing(drawing, global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(240D), global::OfficeIMO.OfficeLength.Points(200D))).ToBytes();

        Assert.Throws<NotSupportedException>(() => Write(Read(0.5D)));
        byte[] ordinary = Write(Read(0D));
        Assert.Equal("B", PdfReadDocument.Open(ordinary).Pages[0].ExtractText().Trim());
    }

    [Fact]
    public void RenderPage_AffineStrokeDrawingPreservesPaintWhenNestedAndClonedWithAnOffset() {
        byte[] pdf = BuildAffineOutlineTextPdf(1D, 0D, 0.5D, 1D, "B", strokeOnly: true,
            strokeWidth: 8D, strokeDash: "[] 0");
        OfficeDrawing source = PdfReadDocument.Open(pdf).Pages[0].ToDrawing();
        var target = new OfficeDrawing(280D, 230D).AddDrawing(source, 15D, 10D).Clone();
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(source, 2D, OfficeColor.White);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(target, 2D, OfficeColor.White);
        int blueInk = 0;
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++) {
            OfficeColor pixel = expected.GetPixel(x, y);
            if (pixel.R < 100 && pixel.G < 100 && pixel.B > 200) blueInk++;
            Assert.Equal(pixel, actual.GetPixel(x + 30, y + 20));
        }
        Assert.True(blueInk > 50, "The source must paint a visible affine blue outline before it is nested.");
    }

    private static byte[] BuildAffineOutlineTextPdf(double a, double b, double c, double d, string text,
        bool clipped = false, string logicalB = "0042", double characterSpacing = 0D, double wordSpacing = 0D,
        bool strokeOnly = false, double strokeWidth = 0.75D, string? strokeDash = null) {
        string Number(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
        int[] widths = new int[35];
        widths[0] = 250; widths[33] = 500; widths[34] = 500;
        string content = (clipped ? "66 100 8 12 re W n " : string.Empty) +
            (strokeOnly ? "0 0 1 RG " + Number(strokeWidth) + " w " + (strokeDash ?? "[2 1] .5") +
                " d 0 J 0 j 10 M " : string.Empty) +
            "BT /F1 20 Tf " + (strokeOnly ? "1 Tr " : string.Empty) +
            Number(characterSpacing) + " Tc " + Number(wordSpacing) + " Tw " +
            $"{Number(a)} {Number(b)} {Number(c)} {Number(d)} 60 100 Tm ({text}) Tj ET";
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            BuildStreamObject(4, "<<", content),
            "5 0 obj\n<< /Type /Font /Subtype /TrueType /BaseFont /AffineOutline /FirstChar 32 /LastChar 66 /Widths [" +
                string.Join(" ", widths) + "] /FontDescriptor 7 0 R /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding " +
                "/Differences [32 /space 65 /A /B] >> /ToUnicode 9 0 R >>\nendobj",
            "6 0 obj\nnull\nendobj",
            "7 0 obj\n<< /Type /FontDescriptor /FontName /AffineOutline /Flags 32 /FontBBox [0 0 600 700] " +
                "/Ascent 800 /Descent -200 /CapHeight 700 /ItalicAngle 0 /StemV 80 /FontFile2 8 0 R >>\nendobj",
            BuildStreamObject(8, "<< /Filter /ASCIIHexDecode", BitConverter.ToString(
                ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 'B')).Replace("-", string.Empty) + ">"),
            BuildStreamObject(9, "<<", "begincmap\n1 begincodespacerange\n<00> <FF>\nendcodespacerange\n" +
                "3 beginbfchar\n<20> <0020>\n<41> <0041>\n<42> <" + logicalB + ">\nendbfchar\nendcmap")
        };
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        var offsets = new long[objects.Length];
        for (int index = 0; index < objects.Length; index++) {
            offsets[index] = output.Position;
            WriteAscii(output, objects[index] + "\n");
        }
        long xref = output.Position;
        WriteAscii(output, "xref\n0 " + (objects.Length + 1).ToString(CultureInfo.InvariantCulture) +
            "\n0000000000 65535 f \n");
        foreach (long offset in offsets)
            WriteAscii(output, offset.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size " + (objects.Length + 1).ToString(CultureInfo.InvariantCulture) +
            " >>\nstartxref\n" + xref.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        return output.ToArray();
    }

    private static OfficePoint[] AffineOutlineCorners(double a, double b, double c, double d,
        double origin, double width) => new[] {
            new OfficePoint(origin, 0D), new OfficePoint(origin + width, 0D),
            new OfficePoint(origin + width, 14D), new OfficePoint(origin, 14D)
        }.Select(point => new OfficePoint(60D + a * point.X + c * point.Y,
            100D + b * point.X + d * point.Y)).ToArray();

    private static void AddAffineOutlinePolygon(OfficeDrawing drawing, OfficePoint[] pdfPoints) {
        OfficePoint[] points = pdfPoints.Select(point => new OfficePoint(point.X, 200D - point.Y)).ToArray();
        var shape = OfficeShape.Polygon(points);
        shape.FillColor = OfficeColor.Black;
        shape.StrokeColor = null;
        drawing.AddShape(shape, points.Min(point => point.X), points.Min(point => point.Y));
    }

    private static void AssertAffineOutlineRaster(string scenario, byte[] pdf, PdfReadPage page, OfficeDrawing expected) {
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(expected, 2D, OfficeColor.White);
        OfficeImageExportResult exported = page.ExportImage(OfficeImageExportFormat.Png,
            new PdfImageExportOptions { TargetDpi = 144D, BackgroundColor = OfficeColor.White });
        Assert.True(OfficePngReader.TryDecode(exported.Bytes, out OfficeRasterImage? decoded));
        OfficeRasterImage actual = decoded!;
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            string directory = Path.Combine(output, "affine-text");
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, scenario + ".pdf"), pdf);
            File.WriteAllBytes(Path.Combine(directory, scenario + "-reference.png"), OfficePngWriter.Encode(reference));
            File.WriteAllBytes(Path.Combine(directory, scenario + "-managed.png"), exported.Bytes);
        }
        AssertAffineRasterPixels(reference, actual);
    }

    private static void AssertAffineOutlineSvgRaster(string scenario, PdfReadPage page, OfficeDrawing expected) {
        OfficeImageExportResult svg = page.ExportImage(OfficeImageExportFormat.Svg, new PdfImageExportOptions());
        Assert.True(OfficeSvgDrawingReader.TryRead(svg.Bytes, out OfficeDrawing? reopened, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(reopened!, 2D, OfficeColor.White);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            string directory = Path.Combine(output, "affine-text");
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, scenario + ".svg"), svg.Bytes);
            File.WriteAllBytes(Path.Combine(directory, scenario + "-svg-managed.png"), OfficePngWriter.Encode(actual));
        }
        AssertAffineRasterPixels(OfficeDrawingRasterRenderer.Render(expected, 2D, OfficeColor.White), actual);
    }

    private static void AssertAffineRasterPixels(OfficeRasterImage reference, OfficeRasterImage actual) {
        Assert.Equal(reference.Width, actual.Width);
        Assert.Equal(reference.Height, actual.Height);
        bool IsInk(OfficeRasterImage image, int x, int y) {
            OfficeColor pixel = image.GetPixel(x, y);
            return pixel.R < 100 || pixel.G < 100 || pixel.B < 100;
        }
        int Missing(OfficeRasterImage wanted, OfficeRasterImage found) {
            int missing = 0;
            for (int y = 0; y < wanted.Height; y++) for (int x = 0; x < wanted.Width; x++) {
                if (!IsInk(wanted, x, y)) continue;
                bool nearby = false;
                for (int dy = -1; dy <= 1 && !nearby; dy++) for (int dx = -1; dx <= 1; dx++) {
                    int nx = x + dx, ny = y + dy;
                    if (nx >= 0 && nx < found.Width && ny >= 0 && ny < found.Height && IsInk(found, nx, ny)) {
                        nearby = true;
                        break;
                    }
                }
                if (!nearby) missing++;
            }
            return missing;
        }
        Assert.True(Enumerable.Range(0, reference.Height).Sum(y =>
            Enumerable.Range(0, reference.Width).Count(x => IsInk(reference, x, y))) > 50,
            "The independently transformed outline must provide visible reference ink.");
        int missing = Missing(reference, actual), extra = Missing(actual, reference);
        Assert.True(missing == 0 && extra == 0,
            $"Affine outline differs beyond one antialiasing pixel: {missing} missing, {extra} extra pixels.");
    }
}
