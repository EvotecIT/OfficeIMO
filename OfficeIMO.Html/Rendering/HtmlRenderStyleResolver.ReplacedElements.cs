using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    // The selected picture source supplies both dimension hints as one owner. A missing
    // hint does not inherit that dimension from the fallback img; CSS still overrides it.
    private IElement ResolveDimensionAttributeSource(IElement element) {
        if (!element.LocalName.Equals("img", StringComparison.OrdinalIgnoreCase)
            || element.ParentElement?.LocalName.Equals("picture", StringComparison.OrdinalIgnoreCase) != true) return element;
        Uri? baseUri = element.Owner is IHtmlDocument document
            ? HtmlDocumentParser.ResolveEffectiveBaseUri(document, _options.BaseUri) : _options.BaseUri;
        HtmlRenderImageSelection selection = HtmlImageSourceResolver.SelectImageForRendering(
            element, baseUri, HtmlResourceUrlPolicy.Create(_options.GetResourceUrlPolicy()), _options);
        return selection.DimensionSource ?? element;
    }

    private void ApplyReplacedElementValues(HtmlComputedStyle computed, double fontSize, HtmlRenderBoxStyle style) {
        var unsupported = new List<string>();
        style.ObjectFit = HtmlCssReplacedElementParser.NormalizeObjectFit(computed.GetValue("object-fit"), out string unsupportedFit);
        if (unsupportedFit.Length > 0) unsupported.Add(unsupportedFit);

        style.ObjectPosition = HtmlCssReplacedElementParser.NormalizeObjectPosition(
            computed.GetValue("object-position"),
            fontSize,
            _rootFontSize,
            _viewportWidth,
            _viewportHeight,
            out string unsupportedPosition);
        if (unsupportedPosition.Length > 0) unsupported.Add(unsupportedPosition);

        style.ApplyEmbeddedImageOrientation = HtmlCssReplacedElementParser.ResolveImageOrientation(
            computed.GetValue("image-orientation"),
            out string unsupportedOrientation);
        if (unsupportedOrientation.Length > 0) unsupported.Add(unsupportedOrientation);

        style.ImageResolutionDpi = HtmlCssReplacedElementParser.ResolveImageResolution(
            computed.GetValue("image-resolution"),
            out string unsupportedResolution);
        if (unsupportedResolution.Length > 0) unsupported.Add(unsupportedResolution);

        if (!HtmlCssReplacedElementParser.TryParseAspectRatio(
                computed.GetValue("aspect-ratio"),
                out style.AspectRatio,
                out style.AspectRatioPrefersIntrinsic,
                out string unsupportedRatio)) {
            style.AspectRatio = null;
            style.AspectRatioPrefersIntrinsic = true;
        }
        if (unsupportedRatio.Length > 0) unsupported.Add(unsupportedRatio);
        style.UnsupportedReplacedElementLayout = string.Join(";", unsupported);
    }
}
