using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextTextAreaAlignmentTests {
    private static double Measure(string? value, double size, string? family, OfficeFontStyle style) =>
        (value?.Length ?? 0) * size / 2D;

    private static OfficeRichTextParagraph Paragraph(string value, OfficeTextAlignment alignment = OfficeTextAlignment.Left,
        OfficeTextPadding? margins = null, OfficeTextParagraphIndent? indent = null) =>
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(value, 10D, OfficeColor.Black) },
            alignment, margins: margins, indent: indent);

    private static OfficeDrawingRichText Text(OfficeDrawing drawing) =>
        Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));

    private static OfficeRichTextBlockLayout Layout(OfficeDrawing drawing, double scale = 1D) {
        OfficeDrawingRichText text = Text(drawing);
        return OfficeDrawingTextLayout.Create(text, (text.Width - text.Padding.Horizontal) * scale,
            (text.Height - text.Padding.Vertical) * scale, Measure, scale);
    }

    [Fact]
    public void ExistingCompiledFriendSignaturesRemainCallable() {
        // Visio and OpenDocument packages already contain member references with these CLR arities.
        var flags = System.Reflection.BindingFlags.NonPublic;
        var builder = typeof(OfficeDrawing).GetMethod("AddRichTextParagraphsCore", flags | System.Reflection.BindingFlags.Instance,
            null, new[] { typeof(System.Collections.Generic.IReadOnlyList<OfficeRichTextParagraph>),
                typeof(double), typeof(double), typeof(double), typeof(double), typeof(OfficeTextVerticalAlignment),
                typeof(bool), typeof(OfficeTextPadding?), typeof(bool) }, null);
        Assert.NotNull(builder);
        var paragraphs = new[] { Paragraph("friend") }; var drawing = new OfficeDrawing(100, 50);
        Assert.Same(drawing, builder!.Invoke(drawing, new object?[] {
            paragraphs, 0D, 0D, 100D, 50D, OfficeTextVerticalAlignment.Top, true, null, false
        }));
        Assert.Equal(OfficeTextAreaAlignment.FullWidth, Text(drawing).TextAreaAlignment);

        var layout = typeof(OfficeDrawingTextLayout).GetMethod("CreateParagraphs", flags | System.Reflection.BindingFlags.Static,
            null, new[] { typeof(System.Collections.Generic.IReadOnlyList<OfficeRichTextParagraph>), typeof(double), typeof(double),
                typeof(Func<string, double, string, OfficeFontStyle, double>), typeof(bool), typeof(double),
                typeof(Func<string, double, string, OfficeFontStyle, OfficeTextPaintBounds>),
                typeof(System.Threading.CancellationToken), typeof(bool), typeof(double) }, null);
        Assert.NotNull(layout);
        var measured = Assert.IsType<OfficeRichTextBlockLayout>(layout!.Invoke(null, new object?[] {
            paragraphs, 100D, 50D, new Func<string?, double, string?, OfficeFontStyle, double>(Measure),
            true, 1D, null, default(System.Threading.CancellationToken), false, 1D
        }));
        Assert.Equal("friend", LineText(Assert.Single(measured.Lines)));
    }

    [Fact]
    public void ExistingParagraphAndRunApisRetainFullWidthPlacement() {
        var paragraphs = new[] { Paragraph("wide\nx", OfficeTextAlignment.Center) };
        var original = new OfficeDrawing(100, 50).AddRichTextParagraphs(paragraphs, 0, 0, 100, 50);
        var explicitDefault = new OfficeDrawing(100, 50).AddRichTextParagraphs(paragraphs, 0, 0, 100, 50,
            OfficeTextAreaAlignment.FullWidth);
        Assert.Equal(OfficeTextAreaAlignment.FullWidth, Text(original).TextAreaAlignment);
        Assert.Equal(new[] { 40D, 47.5D }, Layout(original).Lines.Select(line => line.OffsetX));
        Assert.Equal(OfficeDrawingSvgExporter.ToSvg(original), OfficeDrawingSvgExporter.ToSvg(explicitDefault));
        var runProfile = new OfficeDrawingRichText(paragraphs[0].Runs, 0, 0, 100, 50, OfficeTextAlignment.Center);
        Assert.Equal(OfficeTextAreaAlignment.FullWidth, runProfile.TextAreaAlignment);
    }

    [Theory]
    [InlineData(OfficeTextAreaAlignment.Left, OfficeTextAlignment.Left, 0D, 0D)]
    [InlineData(OfficeTextAreaAlignment.Left, OfficeTextAlignment.Center, 0D, 7.5D)]
    [InlineData(OfficeTextAreaAlignment.Left, OfficeTextAlignment.Right, 0D, 15D)]
    [InlineData(OfficeTextAreaAlignment.Center, OfficeTextAlignment.Left, 40D, 40D)]
    [InlineData(OfficeTextAreaAlignment.Center, OfficeTextAlignment.Center, 40D, 47.5D)]
    [InlineData(OfficeTextAreaAlignment.Center, OfficeTextAlignment.Right, 40D, 55D)]
    [InlineData(OfficeTextAreaAlignment.Right, OfficeTextAlignment.Left, 80D, 80D)]
    [InlineData(OfficeTextAreaAlignment.Right, OfficeTextAlignment.Center, 80D, 87.5D)]
    [InlineData(OfficeTextAreaAlignment.Right, OfficeTextAlignment.Right, 80D, 95D)]
    public void AreaAndParagraphAlignmentRemainIndependent(OfficeTextAreaAlignment area, OfficeTextAlignment paragraph,
        double wideStart, double shortStart) {
        var drawing = new OfficeDrawing(100, 50).AddRichTextParagraphs(new[] { Paragraph("wide\nx", paragraph) },
            0, 0, 100, 50, area, wrapText: false);
        Assert.Equal(OfficeTextAlignment.Left, Text(drawing).Alignment);
        OfficeRichTextBlockLayout layout = Layout(drawing);
        Assert.Equal(new[] { wideStart, shortStart }, layout.Lines.Select(line => line.OffsetX));
        Assert.Equal(new[] { 20D, 5D }, layout.Lines.Select(line => line.Width));
        Assert.False(layout.Clipped);
    }

    [Fact]
    public void ParagraphMarginsContributeToAreaWidthAndFramePaddingIsAppliedOnce() {
        var drawing = new OfficeDrawing(160, 80).AddRichTextParagraphs(new[] {
            Paragraph("wide", OfficeTextAlignment.Right, new OfficeTextPadding(10, 2, 20, 3)),
            Paragraph("x", OfficeTextAlignment.Center, new OfficeTextPadding(5, 4, 15, 6))
        }, 10, 10, 130, 60, OfficeTextAreaAlignment.Center, padding: new OfficeTextPadding(7, 1, 13, 2));
        var layout = Layout(drawing);
        OfficeRichTextLine[] visible = layout.Lines.Where(line => line.Segments.Count > 0).ToArray();
        Assert.Equal(new[] { 40D, 47.5D }, visible.Select(line => line.OffsetX));
        Assert.Equal(39D, layout.Height);
        Assert.Equal(80D, layout.Width);
        Assert.False(layout.Clipped);

        OfficeDrawingRichText text = Text(drawing);
        OfficeRasterCanvas metrics = OfficeDrawingTextLayout.CreateMetrics(drawing);
        double RenderedWidth(string value) => metrics.MeasureText(value, 10, text.Paragraphs[0].Runs[0].FontFamily, OfficeFontStyle.Regular);
        double areaWidth = Math.Max(30D + RenderedWidth("wide"), 20D + RenderedWidth("x"));
        double shift = (110D - areaWidth) / 2D;
        XDocument svg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        Assert.Equal(17D + shift + 10D + areaWidth - 30D - RenderedWidth("wide"), SvgStart(svg, "wide"), 3);
        Assert.Equal(OfficeTextAreaAlignment.Center, text.TextAreaAlignment);
    }

    [Theory]
    [InlineData(OfficeTextAreaAlignment.Center, OfficeTextAlignment.Center, 57.5D, 97.5D)]
    [InlineData(OfficeTextAreaAlignment.Center, OfficeTextAlignment.Right, 57.5D, 137.5D)]
    [InlineData(OfficeTextAreaAlignment.Right, OfficeTextAlignment.Center, 115D, 155D)]
    [InlineData(OfficeTextAreaAlignment.Right, OfficeTextAlignment.Right, 115D, 195D)]
    public void TabsDetermineIntrinsicWidthWhileShortHardLinesUseParagraphAlignment(OfficeTextAreaAlignment area,
        OfficeTextAlignment alignment, double tabbedStart, double shortStart) {
        OfficeRichTextParagraph paragraph = Paragraph("A\tB\nC", alignment).WithTabStops(
            new OfficeTextTabStops(new[] { new OfficeTextTabStop(80D) }).WithParagraphAlignment());
        var drawing = new OfficeDrawing(200, 50).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 200, 50,
            area, wrapText: false);
        OfficeRichTextBlockLayout layout = Layout(drawing);
        Assert.Equal(85D, layout.Lines[0].Width);
        Assert.Equal(tabbedStart, layout.Lines[0].OffsetX);
        Assert.Equal(tabbedStart + 80D, layout.Lines[0].OffsetX + layout.Lines[0].Segments
            .TakeWhile(segment => segment.Text != "B").Sum(segment => segment.Width));
        Assert.Equal(shortStart, layout.Lines[1].OffsetX);
        Assert.False(layout.Clipped);
    }

    [Fact]
    public void IntrinsicPlacementKeepsMeasuredSoftBreaksAdvancesAndHeight() {
        OfficeRichTextParagraph paragraph = Paragraph("a b c", OfficeTextAlignment.Left);
        OfficeDrawing Drawing(OfficeTextAreaAlignment area) => new OfficeDrawing(30, 50)
            .AddRichTextParagraphs(new[] { paragraph }, 0, 0, 30, 50, area);
        double JoinedAdvance(string? value, double size, string? family, OfficeFontStyle style) =>
            (value?.Length ?? 0) * 10D - (value?.Contains("a b") == true ? 5D : 0D);
        OfficeRichTextBlockLayout MeasureDrawing(OfficeDrawing drawing) =>
            OfficeDrawingTextLayout.Create(Text(drawing), 30, 50, JoinedAdvance);
        var original = MeasureDrawing(Drawing(OfficeTextAreaAlignment.FullWidth));
        var intrinsic = MeasureDrawing(Drawing(OfficeTextAreaAlignment.Center));
        Assert.Equal(new[] { "a b", "c" }, intrinsic.Lines.Select(LineText));
        Assert.Equal(original.Lines.Select(LineText), intrinsic.Lines.Select(LineText));
        Assert.Equal(original.Lines.Select(line => line.Width), intrinsic.Lines.Select(line => line.Width));
        Assert.Equal(original.Height, intrinsic.Height);
        Assert.Equal(original.Clipped, intrinsic.Clipped);
        Assert.Equal(2.5D, intrinsic.Lines[0].OffsetX);
        Assert.Equal(2.5D, intrinsic.Lines[1].OffsetX);
    }

    [Fact]
    public void IntrinsicMeasurementDoesNotUseJustifiedAdvances() {
        var drawing = new OfficeDrawing(55, 50).AddRichTextParagraphs(new[] {
            Paragraph("one two three four", OfficeTextAlignment.Justify)
        }, 0, 0, 55, 50, OfficeTextAreaAlignment.Center);
        var layout = Layout(drawing);
        Assert.Equal(new[] { "one two ", "three four" }, layout.Lines.Select(LineText));
        Assert.Equal(new[] { 50D, 50D }, layout.Lines.Select(line => line.Width));
        Assert.All(layout.Lines, line => Assert.Equal(2.5D, line.OffsetX));
    }

    [Theory]
    [InlineData(OfficeTextAreaAlignment.Center, -10D)]
    [InlineData(OfficeTextAreaAlignment.Right, -20D)]
    public void OversizedUnwrappedAreasKeepNegativeCenterAndRightOrigins(OfficeTextAreaAlignment area, double expectedLeft) {
        var drawing = new OfficeDrawing(40, 50).AddRichTextParagraphs(new[] { Paragraph("abcdefghijkl\nx") },
            0, 0, 40, 50, area, wrapText: false);
        var layout = Layout(drawing);
        Assert.Equal(new[] { expectedLeft, expectedLeft }, layout.Lines.Select(line => line.OffsetX));
        Assert.Equal("abcdefghijkl", LineText(layout.Lines[0]));
        Assert.Equal(60D, layout.Lines[0].Width);
        Assert.True(layout.Clipped);
    }

    [Fact]
    public void HeightClippingDoesNotExcludeAHiddenWidestParagraphFromAreaSizing() {
        var drawing = new OfficeDrawing(100, 12).AddRichTextParagraphs(new[] { Paragraph("x"), Paragraph("0123456789") },
            0, 0, 100, 12, OfficeTextAreaAlignment.Center, wrapText: false);
        var layout = Layout(drawing);
        var visible = Assert.Single(layout.Lines);
        Assert.Equal("x", LineText(visible));
        Assert.Equal(25D, visible.OffsetX);
        Assert.Equal(75D, layout.Width);
        Assert.True(layout.Clipped);
    }

    [Fact]
    public void LabelsAndContinuationIndentationRemainInsideTheMovedArea() {
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("1.", 10, OfficeColor.Red), 10,
            followedBy: OfficeTextParagraphLabelFollowedBy.Position, textPosition: 35);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("x\nlong", 10, OfficeColor.Black) }, label,
            margins: new OfficeTextPadding(30, 0, 5, 0), indent: OfficeTextParagraphIndent.Hanging(10));
        var drawing = new OfficeDrawing(100, 50).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 100, 50,
            OfficeTextAreaAlignment.Center, wrapText: false);
        var layout = Layout(drawing);
        Assert.Equal(27.5D, layout.Lines[0].OffsetX);
        Assert.Equal(52.5D, layout.Lines[0].OffsetX + layout.Lines[0].Segments
            .TakeWhile(segment => segment.Text != "x").Sum(segment => segment.Width));
        Assert.Equal(57.5D, layout.Lines[1].OffsetX);
        Assert.Equal(1, layout.Lines.SelectMany(line => line.Segments).Count(segment => segment.Text == "1."));
    }

    [Theory]
    [InlineData(1D)]
    [InlineData(2D)]
    public void FontFittingPrecedesAreaPlacementAndPreservesRenderScale(double scale) {
        var paragraphs = new[] { new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("abcdefghij", 20, OfficeColor.Black) }) };
        OfficeDrawing Drawing(OfficeTextAreaAlignment area) => new OfficeDrawing(40, 12).AddRichTextParagraphsCore(
            paragraphs, 0, 0, 40, 12, OfficeTextVerticalAlignment.Top, false, null, true, area);
        var left = Layout(Drawing(OfficeTextAreaAlignment.Left), scale);
        foreach (OfficeTextAreaAlignment area in new[] { OfficeTextAreaAlignment.Center, OfficeTextAreaAlignment.Right }) {
            var placed = Layout(Drawing(area), scale);
            Assert.False(placed.Clipped);
            OfficeRichTextLine line = Assert.Single(placed.Lines);
            Assert.InRange(line.OffsetX, 0D, 40D * scale);
            Assert.InRange(line.OffsetX + line.Width, 0D, 40D * scale + .01D);
            Assert.Equal(left.Lines[0].FontSize, line.FontSize);
            Assert.Equal(left.Height, placed.Height);
            Assert.InRange(line.FontSize, 6D * scale, 8.01D * scale);
        }
    }

    [Fact]
    public void CloneNestingTintAndFrameTransformsRetainAreaPolicy() {
        var original = new OfficeDrawing(100, 50).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("linked\nx", 10, OfficeColor.Red) { LinkUri = "https://example.invalid/area" } }, OfficeTextAlignment.Right)
        }, 0, 0, 100, 50, OfficeTextAreaAlignment.Center, wrapText: false);
        var clone = original.Clone();
        clone.ApplyColorTint(OfficeColor.Blue);
        var nested = new OfficeDrawing(140, 80).AddDrawing(clone, 20, 10,
            new OfficeImageFrameTransform(20, 70, 40, true, false));
        OfficeDrawingRichText text = Text(nested);
        Assert.Equal(OfficeTextAreaAlignment.Center, text.TextAreaAlignment);
        Assert.Equal(OfficeTextAlignment.Left, text.Alignment);
        Assert.Equal(20D, text.X); Assert.Equal(10D, text.Y);
        Assert.Equal(20D, text.RotationDegrees); Assert.True(text.FlipHorizontal);
        Assert.Equal(OfficeColor.Blue, text.Paragraphs[0].Runs[0].Color);
        Assert.Equal("https://example.invalid/area", text.Paragraphs[0].Runs[0].LinkUri);
        Assert.Equal(OfficeColor.Red, Text(original).Paragraphs[0].Runs[0].Color);
        Assert.Equal(Layout(original).Lines.Select(line => line.OffsetX), Layout(nested).Lines.Select(line => line.OffsetX));
        Assert.Contains("transform=", OfficeDrawingSvgExporter.ToSvg(nested));
    }

    [Theory]
    [InlineData(OfficeTextAreaAlignment.Center, 1D)]
    [InlineData(OfficeTextAreaAlignment.Right, 2D)]
    public void SvgAndRasterUseTheSameLateBoundFontMeasurementForBothLines(OfficeTextAreaAlignment area, double scale) {
        OfficeDrawing Drawing(OfficeTextAreaAlignment placement) => new OfficeDrawing(260, 90).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun("WWWWiiii\n", 12, OfficeColor.Black, fontFamily: "Area Proof"),
                new OfficeRichTextRun("I", 12, OfficeColor.Red, fontFamily: "Area Proof")
            })
        }, 20, 10, 220, 60, placement, wrapText: false, padding: new OfficeTextPadding(10, 0, 14, 0));
        var left = Drawing(OfficeTextAreaAlignment.Left);
        var aligned = Drawing(area);
        double fallback = SvgStart(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(aligned)), "WWWWiiii");
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf"));
        left.Fonts.Add("Area Proof", font); aligned.Fonts.Add("Area Proof", font);
        OfficeRasterCanvas metrics = OfficeDrawingTextLayout.CreateMetrics(aligned);
        double naturalWidth = metrics.MeasureText("WWWWiiii", 12, "Area Proof", OfficeFontStyle.Regular);
        double shift = (196D - naturalWidth) * (area == OfficeTextAreaAlignment.Center ? .5D : 1D);
        XDocument leftSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(left));
        XDocument alignedSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(aligned));
        Assert.NotEqual(fallback, SvgStart(alignedSvg, "WWWWiiii"));
        foreach (string value in new[] { "WWWWiiii", "I" })
            Assert.Equal(shift, SvgStart(alignedSvg, value) - SvgStart(leftSvg, value), 3);
        OfficeRasterImage leftImage = OfficeDrawingRasterRenderer.Render(left, scale, OfficeColor.White);
        OfficeRasterImage alignedImage = OfficeDrawingRasterRenderer.Render(aligned, scale, OfficeColor.White);
        foreach (bool red in new[] { false, true })
            Assert.InRange(FirstInkColumn(alignedImage, red) - FirstInkColumn(leftImage, red), shift * scale - 1D, shift * scale + 1D);
    }

    [Theory]
    [InlineData(OfficeTextAreaAlignment.Center, 1D)]
    [InlineData(OfficeTextAreaAlignment.Right, 2D)]
    public void SvgAndRasterPaintOversizedAreasBeyondTheFramesLeftEdge(OfficeTextAreaAlignment area, double scale) {
        OfficeDrawing Drawing(OfficeTextAreaAlignment placement) {
            var drawing = new OfficeDrawing(400, 100).AddRichTextParagraphs(new[] {
                new OfficeRichTextParagraph(new[] {
                    new OfficeRichTextRun("WWWWWWWW\n", 18, OfficeColor.Black, fontFamily: "Area Proof"),
                    new OfficeRichTextRun("X", 18, OfficeColor.Red, fontFamily: "Area Proof")
                })
            }, 180, 10, 40, 60, placement, wrapText: false, padding: new OfficeTextPadding(2, 0, 4, 0));
            drawing.Fonts.Add("Area Proof", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf")));
            return drawing;
        }
        var left = Drawing(OfficeTextAreaAlignment.Left);
        var aligned = Drawing(area);
        OfficeRasterCanvas metrics = OfficeDrawingTextLayout.CreateMetrics(aligned);
        double naturalWidth = metrics.MeasureText("WWWWWWWW", 18, "Area Proof", OfficeFontStyle.Regular);
        double shift = (34D - naturalWidth) * (area == OfficeTextAreaAlignment.Center ? .5D : 1D);
        Assert.True(shift < 0D);
        var leftSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(left));
        var alignedSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(aligned));
        foreach (string value in new[] { "WWWWWWWW", "X" }) {
            Assert.Equal(shift, SvgStart(alignedSvg, value) - SvgStart(leftSvg, value), 3);
            Assert.True(SvgStart(alignedSvg, value) < 182D);
        }
        OfficeRasterImage leftImage = OfficeDrawingRasterRenderer.Render(left, scale, OfficeColor.White);
        OfficeRasterImage alignedImage = OfficeDrawingRasterRenderer.Render(aligned, scale, OfficeColor.White);
        foreach (bool red in new[] { false, true })
            Assert.InRange(FirstInkColumn(alignedImage, red) - FirstInkColumn(leftImage, red), shift * scale - 1D, shift * scale + 1D);
    }

    private static string LineText(OfficeRichTextLine line) => string.Concat(line.Segments.Select(segment => segment.Text));

    private static double SvgStart(XDocument svg, string value) => double.Parse(Assert.Single(svg.Descendants(),
        element => element.Name.LocalName == "text" && element.Value == value).Attribute("x")!.Value, CultureInfo.InvariantCulture);

    private static int FirstInkColumn(OfficeRasterImage image, bool red) {
        int first = image.Width;
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (red ? pixel.R > 200 && pixel.G < 100 && pixel.B < 100 : pixel.R < 100 && pixel.G < 100 && pixel.B < 100)
                    first = Math.Min(first, x);
            }
        Assert.InRange(first, 0, image.Width - 1);
        return first;
    }
}
