using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Word.Html;

internal partial class HtmlToWordConverter {
    /// <summary>Retains the source's solid HTML canvas when creating a document; insertion keeps the host page.</summary>
    private void ApplyDocumentCanvasBackground(WordDocument document, IHtmlDocument source) {
        IElement? root = source.DocumentElement;
        IElement? body = source.Body;
        string? background = ResolveCanvasBackground(root, out bool rootHasImage);
        // A nontransparent root color owns the canvas. Otherwise HTML propagates the body's color.
        if (background == null && !rootHasImage) background = ResolveCanvasBackground(body, out _);
        if (background != null) document.Background.Color = OfficeColor.Parse("#" + background);
    }

    private string? ResolveCanvasBackground(IElement? element, out bool hasImage) {
        hasImage = false;
        if (element == null) return null;

        // Read the admitted cascade without processing the root as a content element.
        // Custom properties and unrelated root declarations do not describe canvas paint.
        var paint = CollectCssDeclarations(element, inheritedOnly: false)
            .Where(property => property.Key.Equals("background", StringComparison.OrdinalIgnoreCase) ||
                property.Key.StartsWith("background-", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(property => property.Key, property => property.Value.Value, StringComparer.OrdinalIgnoreCase);
        ReportUnsupportedCssDiagnostics(element, paint);
        if (paint.TryGetValue("background-image", out string? image)) {
            hasImage = !string.IsNullOrWhiteSpace(image) && !image!.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                !IsCssWideResetValue(image) && !IsCssInheritanceValue(image);
        }
        paint.TryGetValue("background-color", out string? color);
        CssStyleMapper.CssProperties parsed = CssStyleMapper.ParseStyles("background-color:" + color);
        if (element.ParentElement == null) {
            // The root is not traversed as content; retain its paint for descendant alpha compositing.
            _inlineStyles[element] = parsed;
            InlineStyleParseCount++;
        }
        if (string.IsNullOrEmpty(parsed.BackgroundColor) || parsed.BackgroundColorAlpha is <= 0D) return null;
        return ResolveOpaqueTextBackground(parsed.BackgroundColor!, parsed.BackgroundColorAlpha ?? 1D, null);
    }
}
