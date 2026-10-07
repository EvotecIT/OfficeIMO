using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private static string ResolveTextTransform(IElement element, HtmlComputedStyle computed, string inherited) {
        bool identifier = element.NamespaceUri == "http://www.w3.org/1998/Math/MathML" && element.LocalName == "mi";
        if (identifier && !HasAuthoredValue(computed, "text-transform")) {
            // MathML Core's mi user-agent default is stronger than implicit inheritance.
            return element.GetAttribute("mathvariant") == "normal" && !computed.IsOriginRevertedValue("text-transform")
                ? "none" : "math-auto";
        }
        if (computed.IsResetValue("text-transform")) return "none";
        if (computed.IsInheritedValue("text-transform")) return inherited;
        string value = computed.GetValue("text-transform");
        return string.IsNullOrWhiteSpace(value) ? inherited : value.Trim().ToLowerInvariant();
    }
}
