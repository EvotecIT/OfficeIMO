using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Encloses selected glyph paint in PDF coordinates; whitespace has no painted bounds.</summary>
    private static (double Left, double Bottom, double Right, double Top)? GetGlyphPaintBounds(
        RichSeg segment, string text, double x, double baseline, PdfOptions options,
        Dictionary<byte[], IOfficeBoundedFontProgram> fontCache, PdfOptions? encodingOptions = null,
        PdfTextShowCommand? paintCommand = null, CancellationToken cancellationToken = default) {
        double left = double.PositiveInfinity, bottom = double.PositiveInfinity;
        double right = double.NegativeInfinity, top = double.NegativeInfinity;
        foreach (var glyph in EnumerateGlyphPaintContours(segment, text, x, baseline, options, fontCache,
            encodingOptions, paintCommand, cancellationToken: cancellationToken)) {
            foreach (var contour in glyph.Contours) foreach (OfficePoint point in contour) {
                left = Math.Min(left, point.X); bottom = Math.Min(bottom, point.Y);
                right = Math.Max(right, point.X); top = Math.Max(top, point.Y);
            }
        }
        return double.IsPositiveInfinity(left) ? null : (left, bottom, right, top);
    }

    /// <summary>Uses selected PDF glyphs and PDF-unit positioning, never a separately shaped approximation.</summary>
    /// <remarks>A supplied outline resolver must return private contours for placement. Font-cache eviction also releases its dependent outlines.</remarks>
    private static IEnumerable<(List<List<OfficePoint>> Contours, bool NominalBounds)> EnumerateGlyphPaintContours(
        RichSeg segment, string text, double x, double baseline, PdfOptions options,
        Dictionary<byte[], IOfficeBoundedFontProgram> fontCache, PdfOptions? encodingOptions = null,
        PdfTextShowCommand? paintCommand = null,
        Func<IOfficeBoundedFontProgram, int, double, List<List<OfficePoint>>>? glyphOutline = null,
        Action? fontCacheCleared = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(text)) yield break;
        PdfTextShowCommand command = paintCommand ?? (text == segment.Text && segment.KerningCommand != null ? segment.KerningCommand
            : EncodeTextShowCommand(text, segment.Font, segment.NamedFont, encodingOptions ?? options, segment.FeatureSettings,
                segment.TextDirection, segment.FontMetricScale, preservesPaintedGlyphs: segment.PreservesPaintedGlyphs));
        IOfficeBoundedFontProgram? outline = GetGlyphPaintFont(segment, options, fontCache, fontCacheCleared);
        double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
        double scale = segment.HorizontalTextScaling / 100D;
        double cursor = 0D;
        if (outline == null) {
            for (int index = 0; index < text.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                char value = text[index];
                if (!PdfStandardFontWidths.TryGetHorizontalBounds1000(segment.Font, value, out int left, out int right) ||
                    !PdfStandardFontWidths.TryGetVerticalBounds1000(segment.Font, value, out int bottom, out int top))
                    throw new NotSupportedException("The standard glyph has no qualified nominal paint metrics.");
                // AFM extrema enclose the glyph; the rectangle is not a filled glyph outline.
                if (right > left && top > bottom) {
                    double glyphLeft = x + cursor + left * size / 1000D * scale;
                    double glyphRight = glyphLeft + (right - left) * size / 1000D * scale;
                    double glyphBottom = baseline + bottom * size / 1000D, glyphTop = baseline + top * size / 1000D;
                    yield return (new List<List<OfficePoint>> { new() {
                        new(glyphLeft, glyphBottom), new(glyphRight, glyphBottom), new(glyphRight, glyphTop), new(glyphLeft, glyphTop)
                    } }, true);
                }
                cursor += GetRichGlyphAdvance(segment, text, command, index, options);
            }
            yield break;
        }
        var glyphs = command.VisualGlyphs;
        string hex = command.GlyphHex;
        int count = glyphs?.Count ?? hex.Length / 4;
        for (int index = 0; index < count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int glyphId = glyphs != null ? glyphs[index].GlyphId : GetGlyphPaintId(hex, index * 4);
            double offsetX = glyphs != null ? glyphs[index].OffsetX1000 * size / 1000D : 0D;
            double offsetY = glyphs != null ? glyphs[index].OffsetY1000 * size / 1000D : 0D;
            if (glyphId > 0) {
                // The writer's rounded PDF advances, rather than outline-engine advances, own every origin.
                List<List<OfficePoint>> contours;
                if (glyphOutline != null) contours = glyphOutline(outline, glyphId, size);
                else {
                    var shaped = new OfficeTextShapingResult(new[] { new OfficeShapedGlyph(glyphId, "x", 0) });
                    double ascent = ((IOfficeFontBaselineMetrics)outline).BaselineOffset(size);
                    contours = outline.GetShapedTextContoursBounded("x", shaped, 0D, -ascent, size, 65536, cancellationToken);
                }
                double shear = options.NeedsSyntheticOblique(segment.Font, segment.NamedFont) ? ContentStreamBuilder.SyntheticObliqueShear : 0D;
                foreach (var contour in contours) for (int p = 0; p < contour.Count; p++) {
                    OfficePoint point = contour[p];
                    contour[p] = new OfficePoint(x + cursor + (offsetX + point.X - point.Y * shear) * scale,
                        baseline + offsetY - point.Y);
                }
                yield return (contours, false);
            }
            cursor += GetRichGlyphAdvance(segment, text, command, index, options);
        }
    }

    private static int GetGlyphPaintId(string hex, int offset) {
        int value = 0;
        for (int index = offset; index < offset + 4; index++) {
            char digit = hex[index];
            value = (value << 4) | (digit <= '9' ? digit - '0' : (digit & ~32) - 'A' + 10);
        }
        return value;
    }

    private static IOfficeBoundedFontProgram? GetGlyphPaintFont(RichSeg segment, PdfOptions options,
        Dictionary<byte[], IOfficeBoundedFontProgram> fontCache, Action? fontCacheCleared = null) {
        byte[]? data = null;
        if (segment.NamedFont.HasValue) {
            if (options.TryGetNamedFontProgram(segment.NamedFont.Value, out var named) && named != null) data = named.FontDataForInspection;
            else if (options.TryGetNamedOpenTypeCffFontProgram(segment.NamedFont.Value, out var cff) && cff != null) data = cff.FontDataForInspection;
        }
        if (data == null) {
            if (options.TryGetEmbeddedStandardFontProgram(segment.Font, out var standard) && standard != null) data = standard.FontDataForInspection;
            else if (options.TryGetEmbeddedStandardOpenTypeCffFontProgram(segment.Font, out var cff) && cff != null) data = cff.FontDataForInspection;
        }
        if (data == null) return null;
        if (fontCache.TryGetValue(data, out var cached)) return cached;
        IOfficeBoundedFontProgram font = (IOfficeBoundedFontProgram?)OfficeTrueTypeFont.TryLoad(data)
            ?? OfficeOpenTypeCffFont.TryLoad(data, null, out _) ?? throw new NotSupportedException("The selected PDF font has no inspectable outline program.");
        if (fontCache.Count >= 64) {
            fontCache.Clear();
            fontCacheCleared?.Invoke();
        }
        fontCache.Add(data, font);
        return font;
    }
}
