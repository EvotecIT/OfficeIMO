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

    // These horizontal baseline fixtures prove visible source occurrence and
    // font-specific vertical ink at the saved baseline. Horizontal placement,
    // advance and global whitespace fidelity have separate assertion owners.
    internal static void AssertPdfBackedSvgRetainsBaselineInk(HtmlRenderPage page, string svg) {
        OfficeDrawing drawing = page.CreateDrawing();
        OfficeDrawingText[] painted = DrawingTestTraversal.Elements(drawing).OfType<OfficeDrawingText>()
            .Where(text => !string.IsNullOrWhiteSpace(text.Text)).ToArray();
        Assert.NotEmpty(painted);
        XElement saved = XElement.Parse(svg);
        static string? Label(XElement element) => element.AncestorsAndSelf()
            .Select(ancestor => ancestor.Attribute("aria-label")?.Value).FirstOrDefault(value => value != null);
        XElement[] paths = saved.Descendants().Where(element => element.Name.LocalName == "path"
            && !string.IsNullOrWhiteSpace(Label(element))).ToArray();
        Assert.Equal(painted.Length, paths.Length);
        Assert.Empty(saved.Descendants().Where(element => element.Name.LocalName == "text"
            && !string.IsNullOrWhiteSpace(element.Attribute("aria-label")?.Value ?? element.Value)));
        var matched = new HashSet<XElement>();
        double pointsPerCssPixel = 72D / HtmlRenderOptions.CssPixelsPerInch;
        double[] viewBox = saved.Attribute("viewBox")!.Value.Split(' ')
            .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(4, viewBox.Length);
        Assert.Equal(0D, viewBox[0]);
        Assert.Equal(0D, viewBox[1]);
        Assert.InRange(Math.Abs(viewBox[2] - page.Width * pointsPerCssPixel), 0D, 0.000500001D);
        Assert.InRange(Math.Abs(viewBox[3] - page.Height * pointsPerCssPixel), 0D, 0.000500001D);
        foreach (OfficeDrawingText text in painted) {
            Assert.Equal(OfficeTextAlignment.Left, text.Alignment);
            Assert.Equal(1D, text.BaselineScale);
            Assert.False(text.GlyphTransform.HasValue);
            Assert.False(text.HasFrameTransform);
            OfficePreparedTextShaping? prepared = text.PreparedShaping;
            IOfficeFontProgram? program = prepared?.Program ?? drawing.Fonts.ResolveForText(
                text.RasterText, text.Font.FamilyName, text.Font.Face, text.Font.Size, out _);
            IOfficeBoundedFontProgram bounded = Assert.IsAssignableFrom<IOfficeBoundedFontProgram>(program);
            // Subdivide at the destination font size, as the PDF-backed SVG does.
            // Prepared runs have already been vector-painted by the HTML writer.
            double contourUnitsPerCssPixel = prepared == null ? pointsPerCssPixel : 1D;
            double size = text.Font.Size * contourUnitsPerCssPixel;
            double baseline = program is IOfficeFontBaselineMetrics metrics
                ? metrics.BaselineOffset(size) : program!.LineHeight(size) * .8D;
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
                        .Append(OfficeSvgFormatting.FormatNumber(point.X * outputUnits)).Append(' ')
                        .Append(OfficeSvgFormatting.FormatNumber((text.Y + text.Font.Size + text.BaselineOffset) * pointsPerCssPixel
                            + point.Y * outputUnits)).Append(' ');
                }
                expected.Append('Z');
            }
            Assert.True(expected.Length > 0, "Visible baseline source '" + text.Text.Trim() + "' requires font contour ink.");
            Assert.True(OfficeSvgPathDataParser.TryParse(expected.ToString(), 1_000_000,
                out IReadOnlyList<OfficePathCommand> expectedCommands, out bool expectedLimit) && !expectedLimit,
                "The bounded reference glyph path must retain its command sequence.");
            var comparisons = paths.Where(element => !matched.Contains(element) && Label(element)!.Trim() == text.Text.Trim())
                .Select(element => new { Path = element, Mismatch = BaselineMismatch(element, expectedCommands) }).ToArray();
            var matches = comparisons.Where(item => item.Mismatch == null).ToArray();
            Assert.True(matches.Length > 0,
                "Visible source '" + text.Text.Trim() + "' requires unmatched font-specific baseline ink with "
                + expectedCommands.Count + " exact commands and Y coordinates within 0.001pt; found " + matches.Length
                + " matches among " + comparisons.Length + " trimmed-label candidates.\n"
                + string.Join("\n", comparisons.Take(8).Select((item, index) =>
                    "Candidate " + index + ": " + (item.Mismatch ?? "baseline ink matches"))));
            // Equal labels/baselines can occur more than once. Consume each actual
            // visible slice once instead of inventing a source-order contract.
            Assert.True(matched.Add(matches[0].Path));
        }
        Assert.Equal(paths.Length, matched.Count);

        string? BaselineMismatch(XElement path, IReadOnlyList<OfficePathCommand> reference) {
            if (!OfficeSvgPathDataParser.TryParse(path.Attribute("d")?.Value, 1_000_000,
                    out IReadOnlyList<OfficePathCommand> actual, out bool limit) || limit)
                return "Actual glyph path is empty, invalid or exceeds the bounded command limit.";
            if (reference.Count != actual.Count)
                return "Command count: expected " + reference.Count + ", actual " + actual.Count + ".";
            OfficeTransform transform = OfficeTransform.Identity;
            foreach (XElement ancestor in path.AncestorsAndSelf()) {
                if (!OfficeSvgTransformParser.TryParse(ancestor.Attribute("transform")?.Value, out OfficeTransform current))
                    return "Actual SVG transform is invalid.";
                transform = transform.Then(current);
            }
            OfficePoint[] ink = actual.Where(command => command.Kind != OfficePathCommandKind.Close)
                .Select(command => transform.TransformPoint(command.Point)).ToArray();
            if (ink.Length == 0 || ink.Max(point => point.X) <= ink.Min(point => point.X)
                || ink.Max(point => point.Y) <= ink.Min(point => point.Y))
                return "Actual glyph path has no nonempty two-dimensional ink bounds.";
            const double coordinateTolerance = 0.001D;
            const double floatingPointSlack = 0.000000001D;
            for (int index = 0; index < reference.Count; index++) {
                OfficePathCommand expectedCommand = reference[index], actualCommand = actual[index];
                if (expectedCommand.Kind != actualCommand.Kind)
                    return "Command " + index + ": expected " + expectedCommand.Kind + ", actual " + actualCommand.Kind + ".";
                if (expectedCommand.Kind == OfficePathCommandKind.Close) continue;
                if (expectedCommand.Kind is not (OfficePathCommandKind.MoveTo or OfficePathCommandKind.LineTo))
                    return "Expected glyph command " + index + " is outside this polyline contour fixture.";
                double expectedY = expectedCommand.Point.Y;
                double actualY = transform.TransformPoint(actualCommand.Point).Y;
                double deltaY = actualY - expectedY;
                if (double.IsNaN(deltaY) || double.IsInfinity(deltaY)
                    || Math.Abs(deltaY) > coordinateTolerance + floatingPointSlack)
                    return "First baseline ink mismatch at command " + index + " (" + expectedCommand.Kind + "): expected Y "
                        + Number(expectedY) + ", actual Y " + Number(actualY) + ", delta " + Number(deltaY)
                        + "pt; tolerance " + Number(coordinateTolerance) + "pt.";
            }
            return null;
        }

        static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
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
