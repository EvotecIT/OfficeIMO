using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Uses the selected PDF glyphs and PDF-unit positioning, never a separately shaped text approximation.</summary>
        /// <remarks>Consume each glyph before advancing: embedded outlines use the writer's borrowed placement workspace.</remarks>
        private IEnumerable<(List<List<OfficePoint>> Contours, bool NominalBounds)> DiagnosticGlyphContours(RichSeg segment, string text, double x, double baseline) {
            if (string.IsNullOrWhiteSpace(text)) yield break;
            foreach (var glyph in EnumerateGlyphPaintContours(segment, text, x, baseline, currentOpts,
                _diagnosticFonts, DiagnosticTextOptions, DiagnosticTextCommand(segment, text), cancellationToken: cancellationToken,
                glyphOutline: DiagnosticGlyphOutline, fontCacheCleared: ClearDiagnosticGlyphCache)) yield return glyph;
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
