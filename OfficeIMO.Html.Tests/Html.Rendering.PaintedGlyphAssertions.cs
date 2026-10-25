using OfficeIMO.Drawing;
using OfficeIMO.Html;
using System.Globalization;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    // Scene text remains logical when glyphs are retained. These observations
    // inspect the font-specific glyphs and physical pens carried into painting,
    // rather than requiring one scene visual per scalar or presentation-form text.
    private static PaintedTestGlyph[] ObservePaintedTrueTypeGlyphs(HtmlRenderDocument document,
        IEnumerable<HtmlRenderText> text) {
        var result = new List<PaintedTestGlyph>();
        foreach (HtmlRenderText visual in text) {
            if (visual.PreparedShaping is OfficePreparedTextShaping prepared) {
                OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(prepared.Program);
                OfficeTrueTypeFont.ShapedTextRun run = font.CreateShapedTextRun(prepared.Text, prepared.Result);
                double scale = visual.Font.Size / font.UnitsPerEm;
                double cursor = visual.X;
                for (int index = 0; index < run._glyphs.Length; index++) {
                    OfficeTrueTypeFont.PositionedGlyph glyph = run._glyphs[index];
                    result.Add(new PaintedTestGlyph(visual, font, glyph.GlyphId,
                        prepared.Result.Glyphs[index].UnicodeText,
                        cursor + glyph.OffsetX * scale, visual.Y - glyph.OffsetY * scale));
                    cursor += glyph.AdvanceWidth * scale;
                }
            } else {
                IOfficeFontProgram? resolved = document.Fonts.ResolveForText(visual.Text, visual.Font.FamilyName,
                    visual.FontDescriptor, visual.Font.Size, out _);
                OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(resolved);
                IReadOnlyList<double> origins = font.PositionPortableScalars(visual.Text, visual.Font.Size);
                for (int index = 0; index < visual.Text.Length; index++) {
                    char scalar = visual.Text[index];
                    Assert.False(char.IsSurrogate(scalar)); // These bounded bidi fixtures are BMP-only.
                    Assert.True(font.TryGetGlyphMetrics(scalar, out int glyph, out _));
                    result.Add(new PaintedTestGlyph(visual, font, glyph, scalar.ToString(),
                        visual.X + origins[index], visual.Y));
                }
            }
        }
        return result.OrderBy(item => item.Y).ThenBy(item => item.X).ToArray();
    }

    private static string PhysicalPaintedUnicode(IEnumerable<HtmlRenderText> text) => string.Concat(
        text.OrderBy(item => item.Y).ThenBy(item => item.X).Select(item => item.PreparedShaping == null
            ? item.Text : string.Concat(item.PreparedShaping.Result.Glyphs.Select(glyph => glyph.UnicodeText))));

    private static double PaintBaseline(HtmlRenderText text) => text.Y + text.Font.Size + text.BaselineOffset;

    // These baseline/advance fixtures have ordinary horizontal monochrome slices.
    // Compare each source slice with its actual PDF-backed SVG ink at the exact
    // selected-font pen; a contour label alone does not establish placement.
    internal static void AssertPdfBackedSvgRetainsPaintedGlyphs(HtmlRenderPage page, string svg) {
        // This is the authored HTML scene, not a PDF-readback scene. Its
        // contour/label/pen equality proves the output without a PDF marker.
        OfficeDrawing drawing = page.CreateDrawing();
        OfficeDrawingText[] painted = DrawingTestTraversal.Elements(drawing).OfType<OfficeDrawingText>()
            .Where(text => text.Text.Length > 0).ToArray();
        Assert.NotEmpty(painted);
        XElement saved = XElement.Parse(svg);
        static string? Label(XElement element) => element.AncestorsAndSelf()
            .Select(ancestor => ancestor.Attribute("aria-label")?.Value).FirstOrDefault(value => value != null);
        XElement[] paths = saved.Descendants().Where(element => element.Name.LocalName == "path"
            && Label(element) != null).ToArray();
        XElement[] textNodes = saved.Descendants().Where(element => element.Name.LocalName == "text").ToArray();
        var matched = new HashSet<XElement>();
        double pointsPerCssPixel = 72D / HtmlRenderOptions.CssPixelsPerInch;
        double[] viewBox = saved.Attribute("viewBox")!.Value.Split(' ')
            .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(4, viewBox.Length);
        Assert.Equal(0D, viewBox[0]);
        Assert.Equal(0D, viewBox[1]);
        // The PDF page dimensions are serialized to three decimal places.
        Assert.InRange(Math.Abs(viewBox[2] - page.Width * pointsPerCssPixel), 0D, 0.000500001D);
        Assert.InRange(Math.Abs(viewBox[3] - page.Height * pointsPerCssPixel), 0D, 0.000500001D);
        foreach (OfficeDrawingText text in painted) {
            Assert.Equal(OfficeTextAlignment.Left, text.Alignment);
            Assert.Equal(1D, text.BaselineScale);
            Assert.False(text.GlyphTransform.HasValue);
            Assert.False(text.HasFrameTransform);
            string X(double value) => OfficeSvgFormatting.FormatNumber(value * pointsPerCssPixel);
            OfficePreparedTextShaping? prepared = text.PreparedShaping;
            if (prepared == null) {
                XElement[] scalar = textNodes.Where(element => !matched.Contains(element)
                    && element.Value == text.RasterText && element.Attribute("x")?.Value == X(text.X)
                    && element.Attribute("y")?.Value == X(text.Y + text.Font.Size + text.BaselineOffset)).ToArray();
                if (scalar.Length > 0) {
                    XElement encoded = Assert.Single(scalar);
                    Assert.Equal(text.Text, encoded.Attribute("aria-label")?.Value ?? encoded.Value);
                    Assert.Equal("ltr", encoded.Attribute("direction")?.Value);
                    Assert.Equal("bidi-override", encoded.Attribute("unicode-bidi")?.Value);
                    Assert.Equal(X(text.Font.Size), encoded.Attribute("font-size")?.Value);
                    Assert.True(OfficeFontFamilyParser.Parse(encoded.Attribute("font-family")?.Value)
                        .Contains(text.Font.FamilyName!));
                    Assert.True(text.TextAdvanceWidth.HasValue);
                    Assert.Equal(X(text.TextAdvanceWidth!.Value), encoded.Attribute("textLength")?.Value);
                    Assert.Equal("spacingAndGlyphs", encoded.Attribute("lengthAdjust")?.Value);
                    Assert.True(matched.Add(encoded));
                    continue;
                }
            }
            IOfficeFontProgram? program = prepared?.Program ?? drawing.Fonts.ResolveForText(
                text.RasterText, text.Font.FamilyName, text.Font.Face, text.Font.Size, out _);
            IOfficeBoundedFontProgram bounded = Assert.IsAssignableFrom<IOfficeBoundedFontProgram>(program);
            // Static PDF text is outlined at its destination font size. Adaptive
            // contour subdivision must happen there, before serialization; scaling
            // an already-flattened CSS-pixel contour changes its point sequence.
            // Prepared runs are already vector-painted by the HTML PDF writer.
            double contourUnitsPerCssPixel = prepared == null ? pointsPerCssPixel : 1D;
            double size = text.Font.Size * contourUnitsPerCssPixel;
            double baseline = program is IOfficeFontBaselineMetrics metrics
                ? metrics.BaselineOffset(size) : program!.LineHeight(size) * .8D;
            double advance = prepared?.Measure(size) ?? program!.Measure(text.RasterText, size);
            Assert.True(!double.IsNaN(advance) && !double.IsInfinity(advance) && advance >= 0D);
            double scaleX = text.TextAdvanceWidth.HasValue && advance > 0D
                ? text.TextAdvanceWidth.Value * contourUnitsPerCssPixel / advance : 1D;
            double outputUnits = pointsPerCssPixel / contourUnitsPerCssPixel;
            var contours = prepared != null
                ? prepared.GetContours(0D, -baseline, size, 1_000_000, CancellationToken.None)
                : bounded.GetTextContoursBounded(text.RasterText, 0D, -baseline, size, 1_000_000, CancellationToken.None);
            var expected = new StringBuilder();
            foreach (List<OfficePoint> contour in contours) {
                if (contour.Count < 3) continue;
                for (int index = 0; index < contour.Count; index++) {
                    OfficePoint point = contour[index];
                    expected.Append(index == 0 ? 'M' : 'L')
                        .Append(OfficeSvgFormatting.FormatNumber(text.X * pointsPerCssPixel + point.X * scaleX * outputUnits)).Append(' ')
                        .Append(OfficeSvgFormatting.FormatNumber((text.Y + text.Font.Size + text.BaselineOffset) * pointsPerCssPixel
                            + point.Y * outputUnits)).Append(' ');
                }
                expected.Append('Z');
            }
            if (expected.Length == 0) Assert.True(string.IsNullOrWhiteSpace(text.Text),
                "Only source whitespace may have no contour ink in these fixtures.");
            XElement path = Assert.Single(paths.Where(element => !matched.Contains(element)
                && Label(element) == text.Text && element.Attribute("d")?.Value == expected.ToString()));
            Assert.True(matched.Add(path));
        }
        Assert.Equal(paths.Length + textNodes.Length, matched.Count);
        string logical = string.Concat(saved.Descendants().Where(element => element.Name.LocalName == "text"
            || element.Name.LocalName == "path" && Label(element) != null)
            .Select(element => element.Name.LocalName == "text"
                ? element.Attribute("aria-label")?.Value ?? element.Value : Label(element)));
        Assert.Equal(string.Concat(painted.Select(text => text.Text)), logical);
    }

    private static void AssertSvgRetainsPaintedGlyphs(OfficeDrawing drawing, string svg, string logicalText) {
        static IEnumerable<OfficeDrawingText> TextLeaves(OfficeDrawing source) {
            foreach (OfficeDrawingElement element in source.Elements) {
                if (element is OfficeDrawingText text) yield return text;
                else if (element is OfficeDrawingGroup group) {
                    foreach (OfficeDrawingText nested in TextLeaves(group.Drawing)) yield return nested;
                }
            }
        }
        OfficeDrawingText[] painted = TextLeaves(drawing).Where(item => item.Text.Length > 0).ToArray();
        Assert.NotEmpty(painted);
        Assert.Equal(logicalText, string.Concat(painted.Select(item => item.Text)));
        XElement saved = XElement.Parse(svg);
        static string? LogicalLabel(XElement element) => element.AncestorsAndSelf()
            .Select(ancestor => ancestor.Attribute("aria-label")?.Value).FirstOrDefault(value => value != null);
        XElement[] glyphPaths = saved.Descendants().Where(element => element.Name.LocalName == "path"
            && LogicalLabel(element) != null).ToArray();
        XElement[] positionedText = saved.Descendants().Where(element => element.Name.LocalName == "text").ToArray();
        Assert.True(painted.Length == glyphPaths.Length + positionedText.Length,
            "Every painted source slice must have one encoded path or positioned text slice.\n" + svg);
        string encodedLogical = string.Concat(saved.Descendants()
            .Where(element => element.Name.LocalName == "text"
                || element.Name.LocalName == "path" && LogicalLabel(element) != null)
            .Select(element => element.Name.LocalName == "text"
                ? element.Attribute("aria-label")?.Value ?? element.Value : LogicalLabel(element)));
        Assert.True(encodedLogical == logicalText,
            "The saved slices must retain complete logical text once, through text or contour labels.\n" + svg);
        void EqualAttribute(XElement element, string name, string expected) => Assert.True(
            element.Attribute(name)?.Value == expected,
            "Saved " + name + " must equal " + expected + ".\n" + svg);
        foreach (OfficeDrawingText text in painted) {
            Assert.Equal(OfficeTextAlignment.Left, text.Alignment);
            Assert.Equal(1D, text.BaselineScale);
            Assert.False(text.GlyphTransform.HasValue);
            Assert.False(text.HasFrameTransform);
            if (text.PreparedShaping == null) {
                // An unavailable bounded font face can legitimately retain the
                // already-painted scalar slice as SVG text, without reordering
                // or claiming a font contour that was never available.
                XElement[] matches = positionedText.Where(element => element.Value == text.RasterText
                    && element.Attribute("x")?.Value == OfficeSvgFormatting.FormatNumber(text.X)
                    && element.Attribute("y")?.Value == OfficeSvgFormatting.FormatNumber(text.Y + text.Font.Size + text.BaselineOffset)).ToArray();
                Assert.True(matches.Length == 1, "The scalar glyph and its exact pen must occur once.\n" + svg);
                XElement scalar = matches[0];
                Assert.True(text.PreservesPaintedGlyphs, "Scalar source must carry resolved glyph order.\n" + svg);
                EqualAttribute(scalar, "direction", "ltr");
                EqualAttribute(scalar, "unicode-bidi", "bidi-override");
                EqualAttribute(scalar, "text-anchor", "start");
                EqualAttribute(scalar, "font-size", OfficeSvgFormatting.FormatNumber(text.Font.Size));
                EqualAttribute(scalar, "font-weight", text.Font.Face.Weight.ToString(CultureInfo.InvariantCulture));
                Assert.True(OfficeFontFamilyParser.Parse(scalar.Attribute("font-family")?.Value)
                    .Contains(text.Font.FamilyName!), "The saved glyph must retain its selected font family.\n" + svg);
                Assert.True(text.TextAdvanceWidth.HasValue, "The painted scalar requires its measured advance.\n" + svg);
                EqualAttribute(scalar, "textLength", OfficeSvgFormatting.FormatNumber(text.TextAdvanceWidth!.Value));
                EqualAttribute(scalar, "lengthAdjust", "spacingAndGlyphs");
                Assert.True((scalar.Attribute("aria-label")?.Value ?? scalar.Value) == text.Text,
                    "The saved scalar must retain its logical source.\n" + svg);
                continue;
            }
            OfficePreparedTextShaping prepared = text.PreparedShaping;
            double size = text.Font.Size;
            double baseline = prepared.Program is IOfficeFontBaselineMetrics metrics
                ? metrics.BaselineOffset(size) : prepared.Program.LineHeight(size) * .8D;
            double scaleX = text.TextAdvanceWidth.GetValueOrDefault(prepared.Measure(size)) / prepared.Measure(size);
            var expected = new StringBuilder();
            foreach (List<OfficePoint> contour in prepared.GetContours(0D, -baseline, size, 1_000_000, CancellationToken.None)) {
                if (contour.Count < 3) continue;
                for (int index = 0; index < contour.Count; index++) {
                    OfficePoint point = contour[index];
                    expected.Append(index == 0 ? 'M' : 'L')
                        .Append(OfficeSvgFormatting.FormatNumber(text.X + point.X * scaleX)).Append(' ')
                        .Append(OfficeSvgFormatting.FormatNumber(text.Y + size + text.BaselineOffset + point.Y)).Append(' ');
                }
                expected.Append('Z');
            }
            Assert.NotEqual(string.Empty, expected.ToString());
            // The encoded geometry is compared with the canonical retained font
            // contours, at SVG's defined serialization precision. Different
            // rasterizers need not antialias page edges or outlines identically.
            Assert.True(glyphPaths.Any(path => LogicalLabel(path) == text.Text
                && path.Attribute("d")?.Value == expected.ToString()),
                "Saved glyph contours must equal the retained font geometry at the exact pen.\n" + svg);
        }
        if (positionedText.Length == painted.Length) {
            string encodedPhysical = string.Concat(positionedText.OrderBy(element => double.Parse(element.Attribute("y")!.Value,
                    CultureInfo.InvariantCulture)).ThenBy(element => double.Parse(element.Attribute("x")!.Value,
                    CultureInfo.InvariantCulture)).Select(element => element.Value));
            string paintedPhysical = string.Concat(painted.OrderBy(text => text.Y).ThenBy(text => text.X).Select(text => text.RasterText));
            Assert.True(encodedPhysical == paintedPhysical, "Saved scalar physical order must equal resolved glyph order.\n" + svg);
        }
    }

    private readonly record struct PaintedTestGlyph(HtmlRenderText Visual, OfficeTrueTypeFont Program,
        int GlyphId, string UnicodeText, double X, double Y);
}
