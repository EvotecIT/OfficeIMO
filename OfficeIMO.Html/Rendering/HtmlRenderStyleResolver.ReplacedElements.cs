namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private void ApplyReplacedElementValues(HtmlComputedStyle computed, double fontSize, HtmlRenderBoxStyle style) {
        var unsupported = new List<string>();
        style.ObjectFit = HtmlCssReplacedElementParser.NormalizeObjectFit(computed.GetValue("object-fit"), out string unsupportedFit);
        if (unsupportedFit.Length > 0) unsupported.Add(unsupportedFit);

        style.ObjectPosition = HtmlCssReplacedElementParser.NormalizeObjectPosition(
            computed.GetValue("object-position"),
            fontSize,
            _options.DefaultFontSize,
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
