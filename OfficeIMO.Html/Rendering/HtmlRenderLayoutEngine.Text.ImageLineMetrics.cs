namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Keeps the normal image-line baseline paired with its used CSS line height.
    /// Ordinary paragraph paint and selected-face short lines retain their routes.
    /// </summary>
    private InlineTextLineMetrics ResolveInlineSharedTextMetrics(string text, HtmlRenderBoxStyle style,
        bool hasImage, bool useNormalImageMetrics) {
        if (useNormalImageMetrics && style.Font.Size > 0D && style.LineHeight >= style.Font.Size) {
            HtmlTextFaceMetrics? face = ResolveTextFaceMetrics(text, style);
            if (IsUsableTextFaceMetrics(face)) {
                // CSS line alignment uses pixel-rounded font ascent/descent,
                // leaving the font program's raw metrics and glyph ink intact.
                double ascent = Math.Floor(face!.Value.BaselineOffset + 0.5D);
                double descent = Math.Floor(face.Value.Height - face.Value.BaselineOffset + 0.5D);
                double usedHeight = style.InlineAlignmentLineHeight
                    ?? HtmlRenderCssValues.ResolveInlineAlignmentLineHeight(style.LineHeight, style.Font.Size);
                double leadingUnits = (usedHeight - ascent - descent) * 64D;
                // Divide signed fixed units before flooring the ascent share to
                // a whole pixel. The remaining leading belongs below baseline.
                double ascentLeading = Math.Floor(Math.Truncate(leadingUnits / 2D) / 64D);
                return new InlineTextLineMetrics(usedHeight, ascent + ascentLeading);
            }
        }
        return new InlineTextLineMetrics(style.LineHeight, ResolveInlineTextBaseline(text, style, hasImage));
    }

    // A short-line run can supply the shared image baseline even when the
    // containing paragraph is normal-height. Preserve that qualified route as
    // a whole line rather than changing only its surrounding strut's descent.
    private static bool UsesNormalImageLineMetrics(InlineLine line, HtmlRenderBoxStyle paragraphStyle) =>
        line.HasReplacedImage && paragraphStyle.Font.Size > 0D
        && paragraphStyle.LineHeight >= paragraphStyle.Font.Size
        && !line.Segments.Any(segment => segment.Run.AtomicBlock == null
            && segment.Run.Style.LineHeight < segment.Run.Style.Font.Size)
        && !line.Segments.SelectMany(segment => segment.Run.InlineStrutStyles)
            .Any(style => style.LineHeight < style.Font.Size);

    private readonly struct InlineTextLineMetrics {
        internal InlineTextLineMetrics(double height, double baseline) {
            Height = height;
            Baseline = baseline;
        }

        internal double Height { get; }
        internal double Baseline { get; }
    }
}
