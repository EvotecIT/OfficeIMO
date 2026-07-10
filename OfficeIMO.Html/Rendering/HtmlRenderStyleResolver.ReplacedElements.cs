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
            out string unsupportedPosition);
        if (unsupportedPosition.Length > 0) unsupported.Add(unsupportedPosition);

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
