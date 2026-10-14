namespace OfficeIMO.Html;

/// <summary>A preserved, content-based width to resolve after content measurement.</summary>
internal readonly record struct HtmlRenderIntrinsicWidth(HtmlRenderIntrinsicWidthKind Kind, double? Limit = null);

internal enum HtmlRenderIntrinsicWidthKind {
    MinContent,
    MaxContent,
    FitContent
}
