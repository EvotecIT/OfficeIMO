using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Uses the selected PDF glyphs and PDF-unit positioning, never a separately shaped text approximation.</summary>
        private IEnumerable<(List<List<OfficePoint>> Contours, bool NominalBounds)> DiagnosticGlyphContours(RichSeg segment, string text, double x, double baseline) {
            if (string.IsNullOrWhiteSpace(text)) yield break;
            PdfTextShowCommand command = text == segment.Text && segment.KerningCommand != null ? segment.KerningCommand
                : EncodeTextShowCommand(text, segment.Font, segment.NamedFont, DiagnosticTextOptions, segment.FeatureSettings, segment.TextDirection, segment.FontMetricScale);
            IOfficeBoundedFontProgram? outline = DiagnosticFont(segment);
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
                    if (right > left && top > bottom)
                        yield return (DiagnosticRectangle(x + cursor + left * size / 1000D * scale, baseline + bottom * size / 1000D,
                            (right - left) * size / 1000D * scale, (top - bottom) * size / 1000D), true);
                    double advance = segment.Font is PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique
                        ? 600D : PdfStandardFontWidths.TryGetWidth1000(segment.Font, value, out int width) ? width : 0D;
                    advance += command.SimpleFontKerning?[index] ?? 0;
                    cursor += advance * size / 1000D * scale + segment.CharacterSpacing;
                }
                yield break;
            }
            var glyphs = command.VisualGlyphs;
            string hex = command.GlyphHex;
            int count = glyphs?.Count ?? hex.Length / 4;
            for (int index = 0; index < count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                int glyphId = glyphs != null ? glyphs[index].GlyphId : DiagnosticGlyphId(hex, index * 4);
                double offsetX = glyphs != null ? glyphs[index].OffsetX1000 * size / 1000D : 0D;
                double offsetY = glyphs != null ? glyphs[index].OffsetY1000 * size / 1000D : 0D;
                if (glyphId > 0) {
                    // One glyph avoids accumulating the outline engine's design-unit or tracking
                    // advances. The writer's rounded PDF advances below own every origin.
                    var contours = DiagnosticGlyphOutline(outline, glyphId, size);
                    double shear = currentOpts.NeedsSyntheticOblique(segment.Font, segment.NamedFont) ? ContentStreamBuilder.SyntheticObliqueShear : 0D;
                    foreach (var contour in contours) for (int p = 0; p < contour.Count; p++) {
                        OfficePoint point = contour[p];
                        contour[p] = new OfficePoint(x + cursor + (offsetX + point.X - point.Y * shear) * scale,
                            baseline + offsetY - point.Y);
                    }
                    yield return (contours, false);
                }
                double advance;
                if (glyphs != null) advance = glyphs[index].AdvanceWidth1000;
                else if (segment.NamedFont.HasValue && currentOpts.TryGetNamedFontProgram(segment.NamedFont.Value, out var named) && named != null)
                    advance = named.GetGlyphWidth1000(glyphId);
                else if (currentOpts.TryGetEmbeddedStandardFontProgram(segment.Font, out var standard) && standard != null)
                    advance = standard.GetGlyphWidth1000(glyphId);
                else if (segment.NamedFont.HasValue && currentOpts.TryGetNamedOpenTypeCffFontProgram(segment.NamedFont.Value, out var namedCff) && namedCff != null)
                    advance = namedCff.GetGlyphWidth1000(glyphId);
                else if (currentOpts.TryGetEmbeddedStandardOpenTypeCffFontProgram(segment.Font, out var standardCff) && standardCff != null)
                    advance = standardCff.GetGlyphWidth1000(glyphId);
                else throw new InvalidOperationException("The selected embedded glyph has no PDF advance metrics.");
                double tracking = command.TrackingBoundaries?[index] == true ? GetIntrinsicGlyphTrackingAdvance(command, size) : 0D;
                cursor += (advance * size / 1000D + tracking) * scale + segment.CharacterSpacing;
            }
        }

        private static int DiagnosticGlyphId(string hex, int offset) {
            int value = 0;
            for (int index = offset; index < offset + 4; index++) {
                char digit = hex[index];
                value = (value << 4) | (digit <= '9' ? digit - '0' : (digit & ~32) - 'A' + 10);
            }
            return value;
        }

        private IOfficeBoundedFontProgram? DiagnosticFont(RichSeg segment) {
            byte[]? data = null;
            if (segment.NamedFont.HasValue) {
                if (currentOpts.TryGetNamedFontProgram(segment.NamedFont.Value, out var named) && named != null) data = named.FontDataForInspection;
                else if (currentOpts.TryGetNamedOpenTypeCffFontProgram(segment.NamedFont.Value, out var cff) && cff != null) data = cff.FontDataForInspection;
            }
            if (data == null) {
                if (currentOpts.TryGetEmbeddedStandardFontProgram(segment.Font, out var standard) && standard != null) data = standard.FontDataForInspection;
                else if (currentOpts.TryGetEmbeddedStandardOpenTypeCffFontProgram(segment.Font, out var cff) && cff != null) data = cff.FontDataForInspection;
            }
            if (data == null) return null;
            if (_diagnosticFonts.TryGetValue(data, out var cached)) return cached;
            IOfficeBoundedFontProgram font = (IOfficeBoundedFontProgram?)OfficeTrueTypeFont.TryLoad(data)
                ?? OfficeOpenTypeCffFont.TryLoad(data, null, out _) ?? throw new NotSupportedException("The selected PDF font has no inspectable outline program.");
            if (_diagnosticFonts.Count >= 64) {
                _diagnosticFonts.Clear();
                // Outline keys retain their parsed font. Release them with the font cache.
                _diagnosticGlyphContours.Clear();
                _diagnosticGlyphPointCount = 0;
            }
            _diagnosticFonts.Add(data, font);
            return font;
        }

        private IEnumerable<List<List<OfficePoint>>> DiagnosticDecorationContours(double left, double right, double y, OfficeTextDecorationStyle style) {
            const double thickness = .5D;
            if (style == OfficeTextDecorationStyle.Double) {
                double separation = Math.Max(thickness * 1.8D, .8D);
                foreach (var part in DiagnosticDecorationContours(left, right, y - separation / 2D, OfficeTextDecorationStyle.Single)) yield return part;
                foreach (var part in DiagnosticDecorationContours(left, right, y + separation / 2D, OfficeTextDecorationStyle.Single)) yield return part;
                yield break;
            }
            var points = new List<OfficePoint> { new(left, y) };
            if (style == OfficeTextDecorationStyle.Wavy) {
                double step = Math.Max(thickness * 5D, 2D), amplitude = Math.Max(thickness * 1.5D, .75D);
                bool raised = true;
                for (double cursor = left; cursor < right;) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (points.Count >= 65536 || cursor + step <= cursor) throw new NotSupportedException("Decoration paint exceeds its bounded contour inspection limit.");
                    cursor = Math.Min(right, cursor + step);
                    points.Add(new OfficePoint(cursor, y + (raised ? amplitude : -amplitude))); raised = !raised;
                }
            } else points.Add(new OfficePoint(right, y));
            double[]? dash = style == OfficeTextDecorationStyle.Dashed ? new[] { 2.5D, 1.5D }
                : style == OfficeTextDecorationStyle.Dotted ? new[] { .5D, 1.5D } : null;
            var contours = new[] { new OfficeFlattenedPathContour(points, false) };
            yield return OfficeStrokeGeometry.Create(contours, thickness,
                style == OfficeTextDecorationStyle.Dotted ? OfficeStrokeLineCap.Round : OfficeStrokeLineCap.Butt,
                style == OfficeTextDecorationStyle.Wavy ? OfficeStrokeLineJoin.Round : OfficeStrokeLineJoin.Miter,
                10D, dash, 0D, 4D, double.NegativeInfinity, double.NegativeInfinity, double.PositiveInfinity, double.PositiveInfinity, cancellationToken);
        }
    }
}
