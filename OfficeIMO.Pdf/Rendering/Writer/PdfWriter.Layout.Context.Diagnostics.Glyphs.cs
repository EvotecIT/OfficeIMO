using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Uses the selected PDF glyphs and PDF-unit positioning, never a separately shaped text approximation.</summary>
        /// <remarks>Callers enumerate immediately and consume each glyph before advancing; embedded outlines use the writer's borrowed placement workspace.</remarks>
        private IEnumerable<(List<List<OfficePoint>> Contours, bool NominalBounds)> DiagnosticGlyphContours(RichSeg segment, string text, double x, double baseline,
            Func<DiagnosticPaintBounds, bool>? contained = null) {
            if (string.IsNullOrWhiteSpace(text)) return Array.Empty<(List<List<OfficePoint>> Contours, bool NominalBounds)>();
            List<List<OfficePoint>>? ResolveOutline(IOfficeBoundedFontProgram font, int glyph, double size, GlyphPaintPlacement placement) =>
                DiagnosticGlyphOutline(font, glyph, size, placement, contained);
            PdfCreateOptions encodingOptions = DiagnosticTextOptions;
            PdfTextShowCommand command = DiagnosticTextCommand(segment, text);
            if (contained != null && SegmentGlyphInkContained(segment, text, x, baseline, command, contained))
                return Array.Empty<(List<List<OfficePoint>> Contours, bool NominalBounds)>();
            return EnumerateGlyphPaintContours(segment, text, x, baseline, currentOpts,
                _diagnosticFonts, encodingOptions, command, cancellationToken: cancellationToken,
                glyphOutline: ResolveOutline, fontCacheCleared: ClearDiagnosticGlyphCache, nominalContained: contained);
        }

        /// <summary>
        /// Proves for a whole segment what <see cref="DiagnosticGlyphOutline"/> proves glyph by glyph: that every
        /// inked glyph's placed outline extrema lie inside the paint clips. It places each glyph exactly as the
        /// glyph walk does and asks <paramref name="contained"/> once about the union of those boxes. The clip test
        /// maps both axes with positive scales, so the union passes exactly when every glyph would pass. Any glyph
        /// the per-glyph route would not settle by its extrema fails the proof, and the caller walks the glyphs.
        /// </summary>
        private bool SegmentGlyphInkContained(RichSeg segment, string text, double x, double baseline,
            PdfTextShowCommand command, Func<DiagnosticPaintBounds, bool> contained) {
            text = WithoutOptionalBreaks(text);
            IOfficeBoundedFontProgram? outline = GetGlyphPaintFont(segment, currentOpts, _diagnosticFonts, ClearDiagnosticGlyphCache);
            if (outline is not OfficeTrueTypeFont) return false;
            double size = EffectiveRichFontSize(segment.FontSize, segment.Baseline);
            double scale = segment.HorizontalTextScaling / 100D;
            if (!(scale > 0D) || double.IsInfinity(scale)) return false;
            var advances = new RichGlyphAdvanceReader(segment, text, command, currentOpts);
            var glyphs = command.VisualGlyphs;
            string hex = command.GlyphHex;
            int count = glyphs?.Count ?? hex.Length / 4;
            bool stylesResolved = false, syntheticBold = false;
            double widen = 0D, cursor = 0D;
            double left = double.PositiveInfinity, top = double.PositiveInfinity;
            double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
            bool inked = false;
            for (int index = 0; index < count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                int glyphId = glyphs != null ? glyphs[index].GlyphId : GetGlyphPaintId(hex, index * 4);
                if (glyphId > 0) {
                    if (!stylesResolved) {
                        // A sheared placement is not monotone in x, so its glyphs keep the outline walk.
                        if (currentOpts.NeedsSyntheticOblique(segment.Font, segment.NamedFont)) return false;
                        syntheticBold = currentOpts.NeedsSyntheticBold(segment.Font, segment.NamedFont);
                        widen = syntheticBold ? ContentStreamBuilder.SyntheticBoldStrokeWidth(size) / 2D * (1D + 1E-9D) + 1E-9D : 0D;
                        stylesResolved = true;
                    }
                    double offsetX = glyphs != null ? glyphs[index].OffsetX1000 * size / 1000D : 0D;
                    double offsetY = glyphs != null ? glyphs[index].OffsetY1000 * size / 1000D : 0D;
                    var placement = new GlyphPaintPlacement(x + cursor, baseline, offsetX, offsetY, scale, 0D, syntheticBold);
                    var cached = GetDiagnosticGlyphOutline(outline, glyphId, size);
                    if (!HasNoOutlinePoints(cached.Contours)) {
                        if (!cached.Bounds.IsFinite) return false;
                        OfficePoint first = placement.TransformPoint(new OfficePoint(cached.Bounds.Left, cached.Bounds.Bottom));
                        OfficePoint second = placement.TransformPoint(new OfficePoint(cached.Bounds.Right, cached.Bounds.Top));
                        var bounds = new DiagnosticPaintBounds(first.X - widen, first.Y - widen, second.X + widen, second.Y + widen);
                        if (!bounds.IsFinite) return false;
                        left = Math.Min(left, bounds.Left);
                        top = Math.Min(top, bounds.Top);
                        right = Math.Max(right, bounds.Right);
                        bottom = Math.Max(bottom, bounds.Bottom);
                        inked = true;
                    }
                }
                cursor += advances.Get(index);
            }
            return !inked || contained(new DiagnosticPaintBounds(left, top, right, bottom));
        }

        private const int DiagnosticTextCommandLimit = 4096;
        private readonly Dictionary<(string Text, PdfStandardFont Font, PdfNamedFontFace? NamedFont, OfficeTextFeatureSettings Features,
            OfficeTextDirection Direction, double MetricScale, bool PreservesPaintedGlyphs), PdfTextShowCommand> _diagnosticTextCommands = new();
        private PdfCreateOptions? _diagnosticTextCommandOptions;

        /// <summary>
        /// Encodes inspected text with a measurement snapshot once per distinct run. Encoding with a
        /// snapshot has no effect on the output document, so equal inputs give an equal command.
        /// </summary>
        private PdfTextShowCommand DiagnosticTextCommand(RichSeg segment, string text) {
            if (text == segment.Text && segment.KerningCommand != null) return segment.KerningCommand;
            PdfCreateOptions options = DiagnosticTextOptions;
            if (_diagnosticTextCommandOptions == null || !EncodesTextAlike(_diagnosticTextCommandOptions, options)
                || _diagnosticTextCommands.Count >= DiagnosticTextCommandLimit) {
                _diagnosticTextCommands.Clear();
                _diagnosticTextCommandOptions = options;
            }
            var key = (text, segment.Font, segment.NamedFont, segment.FeatureSettings, segment.TextDirection,
                segment.FontMetricScale, segment.PreservesPaintedGlyphs);
            if (!_diagnosticTextCommands.TryGetValue(key, out PdfTextShowCommand? command)) {
                command = EncodeTextShowCommand(text, segment.Font, segment.NamedFont, options, segment.FeatureSettings,
                    segment.TextDirection, segment.FontMetricScale, preservesPaintedGlyphs: segment.PreservesPaintedGlyphs);
                _diagnosticTextCommands.Add(key, command);
            }
            return command;
        }

        // Pages and frames get their own options; a snapshot with the same fonts and shaping settings encodes text the same way.
        private static bool EncodesTextAlike(PdfCreateOptions cached, PdfCreateOptions current) =>
            ReferenceEquals(cached, current)
            || cached.HasSameMeasurementFonts(current)
            && cached.TextShapingModeSnapshot == current.TextShapingModeSnapshot
            && ReferenceEquals(cached.TextShapingProviderSnapshot, current.TextShapingProviderSnapshot)
            && string.Equals(cached.EffectiveTextShapingLanguage, current.EffectiveTextShapingLanguage, StringComparison.Ordinal)
            && cached.UseFontUnitTextLayout == current.UseFontUnitTextLayout
            && ReferenceEquals(cached.DefaultTextFeatureSettings, current.DefaultTextFeatureSettings);

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
                double step = WavyDecorationStep(thickness), amplitude = WavyDecorationAmplitude(thickness);
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
