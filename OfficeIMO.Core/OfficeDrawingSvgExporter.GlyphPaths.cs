using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    // Painted runs need no reshaping. The explicit option additionally permits
    // simple logical runs shaped by the shared managed/host shaping provider.
    // Unsupported layouts retain text and their registered fonts.
    private static bool TryAppendGlyphPaths(StringBuilder output, OfficeDrawingText text, OfficeRasterCanvas metrics, bool textAsPaths) {
        OfficePreparedTextShaping? prepared = text.PreparedShaping;
        if ((prepared == null && !text.PreservesPaintedGlyphs && !textAsPaths) || text.WrapText || text.ShrinkToFit || text.StackedText || text.HasPadding ||
            text.VerticalAlignment != OfficeTextVerticalAlignment.Top || text.TextDirection == OfficeTextDirection.TopToBottom ||
            text.Font.IsUnderline || text.Font.IsStrikethrough || text.UnderlineStyle != OfficeTextDecorationStyle.None ||
            text.StrikethroughStyle != OfficeTextDecorationStyle.None || text.FontMetricScale != 1D ||
            text.RasterText.IndexOfAny(new[] { '\r', '\n' }) >= 0 ||
            prepared == null && text.FeatureSettings is { IsDefault: false } ||
            !string.IsNullOrEmpty(text.FontPalette) && text.FontPalette != "normal") {
            if (prepared != null) throw new NotSupportedException(
                "Prepared SVG glyphs require one unwrapped horizontal monochrome slice without decorations or specialized text geometry.");
            return false;
        }
        bool logical = prepared == null && !text.PreservesPaintedGlyphs && !text.GlyphTransform.HasValue;
        // Bound the first opt-in logical-text slice to printable ASCII. Complex
        // scripts, controls and bidi remain with the browser's shaping stack.
        if (logical && (text.TextDirection == OfficeTextDirection.RightToLeft ||
            text.RasterText.Any(character => character < ' ' || character > '~'))) return false;
        double sourceSize = text.Font.Size > 0D ? text.Font.Size : 10D;
        double size = sourceSize * text.BaselineScale;
        IOfficeBoundedFontProgram font;
        if (prepared != null) font = (IOfficeBoundedFontProgram)prepared.Program;
        else {
            if (metrics.Fonts == null) return false;
            // The normal route resolves fonts per complete text element. Retained glyph ids
            // already belong to their pinned program and must bypass that resolution.
            IReadOnlyList<OfficeFontFallbackRun> fallbackRuns = metrics.Fonts.PlanFallbackRuns(
                text.RasterText, text.Font.FamilyName, text.Font.Face);
            if (fallbackRuns.Count != 1) return false;
            if (!metrics.Fonts.TryResolveFaceForText(text.RasterText, fallbackRuns[0].FamilyRequest,
                    text.Font.Face, size, out OfficeFontFace? face) || face == null ||
                face.Program is not IOfficeBoundedFontProgram resolved ||
                face.Style != OfficeFontFace.NormalizeStyle(text.Font.Style)) return false;
            font = resolved;
        }
        var token = metrics.CancellationToken;
        OfficeTextShapingResult? shaped = prepared?.Result;
        if (logical) {
            IOfficeTextShapingProvider provider = metrics.TextShapingProvider ?? OfficeManagedTextShapingProvider.Instance;
            token.ThrowIfCancellationRequested();
            shaped = provider.ShapeText(new OfficeTextShapingRequest(text.RasterText,
                font.DisplayName ?? string.Empty,
                font is OfficeTrueTypeFont trueType ? trueType.FontDataForShaping : font.GetFontDataForShaping(),
                font.IsOpenTypeCff, font.UnitsPerEm, OfficeTextDirection.LeftToRight,
                metrics.TextShapingLanguage, token, font.CollectionIndex,
                (font as IOfficeVariableFontProgram)?.VariationCoordinatesForShaping,
                cloneFontData: false, fontProgramCacheKey: font, featureSettings: text.FeatureSettings,
                applyDefaultLatinLigatures: true));
            token.ThrowIfCancellationRequested();
            if (shaped == null || shaped.Direction is OfficeTextDirection.RightToLeft or OfficeTextDirection.TopToBottom) return false;
        }
        // Color glyphs need their layered/bitmap paint, rather than monochrome contours.
        if (font is IOfficeColorFontProgram colorFont) {
            if (shaped != null && shaped.Glyphs.Any(glyph => colorFont.HasColorGlyph(glyph.GlyphId))) return false;
            for (int index = 0; index < text.RasterText.Length; index++) {
                int scalar = char.IsHighSurrogate(text.RasterText[index]) && index + 1 < text.RasterText.Length &&
                    char.IsLowSurrogate(text.RasterText[index + 1])
                    ? char.ConvertToUtf32(text.RasterText[index], text.RasterText[++index]) : text.RasterText[index];
                if (font.TryGetGlyphMetrics(scalar, out int glyph, out _) && colorFont.HasColorGlyph(glyph)) return false;
            }
        }
        double advance = prepared?.Measure(size) ?? (shaped == null
            ? font.Measure(text.RasterText, size) : font.MeasureShapedText(text.RasterText, shaped, size));
        if (advance < 0D || advance == 0D && prepared == null && !text.PreservesPaintedGlyphs ||
            double.IsNaN(advance) || double.IsInfinity(advance)) return false;
        double scaleX = text.TextAdvanceWidth.HasValue && advance > 0D ? text.TextAdvanceWidth.Value / advance : 1D;
        double width = advance * scaleX;
        double x = text.X + (text.Alignment == OfficeTextAlignment.Center ? (text.Width - width) / 2D :
            text.Alignment == OfficeTextAlignment.Right ? text.Width - width : 0D);
        double y = text.Y + sourceSize + text.BaselineOffset;
        double baseline = font is IOfficeFontBaselineMetrics baselineMetrics
            ? baselineMetrics.BaselineOffset(size) : font.LineHeight(size) * .8D;
        const int MaximumOutlinePoints = 1_000_000;
        List<List<OfficePoint>> contours;
        if (prepared != null) {
            contours = prepared.GetContours(0D, -baseline, size, MaximumOutlinePoints, token,
                new OfficeCffOperationBudget());
        } else if (shaped == null) {
            contours = font is IOfficeCffBoundedFontProgram cff
                ? cff.GetTextContoursBounded(text.RasterText, 0D, -baseline, size, MaximumOutlinePoints, token,
                    new OfficeCffOperationBudget())
                : font.GetTextContoursBounded(text.RasterText, 0D, -baseline, size, MaximumOutlinePoints, token);
        } else {
            contours = font is IOfficeCffBoundedFontProgram cff
                ? cff.GetShapedTextContoursBounded(text.RasterText, shaped, 0D, -baseline, size, MaximumOutlinePoints, token,
                    new OfficeCffOperationBudget())
                : font.GetShapedTextContoursBounded(text.RasterText, shaped, 0D, -baseline, size, MaximumOutlinePoints, token);
        }
        output.Append("<g").AppendAttribute("role", "img").AppendAttribute("aria-label", text.Text);
        string? transform = OfficeSvgFormatting.FormatImageFrameTransform(text.CreateFrameTransform());
        if (transform != null) output.AppendAttribute("transform", transform);
        output.Append("><path d=\"");
        foreach (List<OfficePoint> contour in contours) {
            token.ThrowIfCancellationRequested();
            if (contour.Count < 3) continue;
            for (int index = 0; index < contour.Count; index++) {
                token.ThrowIfCancellationRequested();
                OfficePoint point = contour[index];
                var placed = new OfficePoint(x + point.X * scaleX, y + point.Y);
                if (text.GlyphTransform.HasValue) placed = text.GlyphTransform.Value.TransformPoint(placed);
                output.Append(index == 0 ? 'M' : 'L').Append(Format(placed.X))
                    .Append(' ').Append(Format(placed.Y)).Append(' ');
            }
            output.Append('Z');
        }
        output.Append('"').AppendAttribute("fill-rule", "nonzero");
        if (text.FillSuppressed) output.AppendAttribute("fill", "none");
        else output.AppendPaintAttribute("fill", text.Color ?? OfficeColor.Black);
        if (text.OutlineStrokeColor is OfficeColor stroke) AppendTextStrokeAttributes(output, stroke, text.OutlineStrokeWidth, text.OutlineStrokeDashPattern, text.OutlineStrokeDashOffset, text.OutlineStrokeLineCap, text.OutlineStrokeLineJoin, text.OutlineStrokeMiterLimit);
        output.Append("/></g>");
        return true;
    }

    private static void AppendTextStrokeAttributes(StringBuilder output, OfficeColor color, double width,
        IReadOnlyList<double>? dashPattern, double dashOffset, OfficeStrokeLineCap cap, OfficeStrokeLineJoin join, double miterLimit) {
        output.AppendPaintAttribute("stroke", color)
            .AppendNumberAttribute("stroke-width", width > 0D ? width : 1D)
            .AppendStrokeLineCapAttribute(cap)
            .AppendStrokeLineJoinAttribute(join)
            .AppendNumberAttribute("stroke-miterlimit", miterLimit);
        if (dashPattern is { Count: > 0 }) {
            output.Append(" stroke-dasharray=\"");
            for (int index = 0; index < dashPattern.Count; index++) {
                if (index > 0) output.Append(' ');
                output.Append(Format(dashPattern[index]));
            }
            output.Append('"');
        }
        if (dashOffset != 0D) output.AppendNumberAttribute("stroke-dashoffset", dashOffset);
        if (width <= 0D) output.AppendAttribute("vector-effect", "non-scaling-stroke");
    }

    // Keep stroke-only text even when the fill is transparent: the text formatter
    // normally omits transparent runs before the SVG paint attributes can be added.
    private static OfficeColor ResolveSvgTextFill(OfficeDrawingText text) =>
        text.OutlineStrokeColor.HasValue && (text.Color ?? OfficeColor.Black).A == 0
            ? OfficeColor.Black : text.Color ?? OfficeColor.Black;

    private static void AppendStrokedSvgText(StringBuilder output, string fragment, OfficeDrawingText text, CancellationToken token) {
        var root = XElement.Parse("<g>" + fragment + "</g>", LoadOptions.PreserveWhitespace);
        var attributes = new StringBuilder("<g");
        AppendTextStrokeAttributes(attributes, text.OutlineStrokeColor!.Value, text.OutlineStrokeWidth, text.OutlineStrokeDashPattern, text.OutlineStrokeDashOffset, text.OutlineStrokeLineCap, text.OutlineStrokeLineJoin, text.OutlineStrokeMiterLimit);
        var paint = XElement.Parse(attributes.Append("/>").ToString());
        foreach (XElement element in root.Descendants().Where(e => e.Name.LocalName is "text" or "tspan")) {
            token.ThrowIfCancellationRequested();
            foreach (XAttribute attribute in paint.Attributes()) element.SetAttributeValue(attribute.Name, attribute.Value);
            if (text.FillSuppressed) {
                element.SetAttributeValue("fill", "none");
                element.Attribute("fill-opacity")?.Remove();
            } else {
                OfficeColor fill = text.Color ?? OfficeColor.Black;
                element.SetAttributeValue("fill", "#" + fill.ToRgbHex());
                element.SetAttributeValue("fill-opacity", (fill.A / 255D).ToString("R", System.Globalization.CultureInfo.InvariantCulture));
            }
        }
        foreach (XNode node in root.Nodes()) output.Append(node.ToString(SaveOptions.DisableFormatting));
    }

    private static void CollectSvgTextFontFamilies(string fragment, ISet<string> families, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (fragment.Length == 0) return;
        var root = XElement.Parse("<g>" + fragment + "</g>", LoadOptions.PreserveWhitespace);
        foreach (XElement element in root.Descendants().Where(e => e.Name.LocalName is "text" or "tspan")) {
            token.ThrowIfCancellationRequested();
            string? value = element.Attribute("font-family")?.Value;
            if (value == null) continue;
            foreach (string family in OfficeFontFamilyParser.Parse(value)) families.Add(family);
        }
    }
}
