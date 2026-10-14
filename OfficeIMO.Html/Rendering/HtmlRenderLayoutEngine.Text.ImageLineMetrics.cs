namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Keeps the normal image-line baseline paired with its used CSS line height.
    /// Selected-face short image lines retain their qualified route.
    /// </summary>
    private InlineTextLineMetrics ResolveInlineSharedTextMetrics(string text, HtmlRenderBoxStyle style,
        bool hasImage, bool useNormalImageMetrics) {
        if (useNormalImageMetrics && style.Font.Size > 0D && style.LineHeight >= style.Font.Size
            && TryResolveNormalInlineTextMetrics(style, ResolveTextFaceMetrics(text, style), out InlineTextLineMetrics normal))
            return normal;
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

}
