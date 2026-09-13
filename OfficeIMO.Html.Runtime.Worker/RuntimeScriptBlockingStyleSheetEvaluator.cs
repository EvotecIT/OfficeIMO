using AngleSharp.Css;
using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeScriptBlockingStyleSheetEvaluator(HtmlScriptRequest request) : IScriptBlockingStyleSheetEvaluator {
    public bool Matches(string mediaText) => HtmlComputedStyleEngine.IsApplicableMedia(
        mediaText,
        HtmlCssMediaContext.Screen,
        request.ViewportWidth,
        request.ViewportHeight);
}
