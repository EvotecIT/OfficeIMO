using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlContentSafety {
    private static bool IsGeneratedContentAttributeReference(
        IElement element,
        HtmlComputedStyleSet styleSet,
        string attributeName) =>
        PseudoContentReferencesAttribute(element, styleSet, HtmlPseudoElementKind.Before, attributeName)
        || PseudoContentReferencesAttribute(element, styleSet, HtmlPseudoElementKind.After, attributeName)
        || PseudoContentReferencesAttribute(element, styleSet, HtmlPseudoElementKind.Marker, attributeName)
        || PseudoContentReferencesAttribute(element, styleSet, HtmlPseudoElementKind.FootnoteCall, attributeName)
        || PseudoContentReferencesAttribute(element, styleSet, HtmlPseudoElementKind.FootnoteMarker, attributeName);

    private static bool PseudoContentReferencesAttribute(
        IElement element,
        HtmlComputedStyleSet styleSet,
        HtmlPseudoElementKind kind,
        string attributeName) =>
        styleSet.TryGetPseudoStyle(element, kind, out HtmlComputedStyle style)
        && HtmlGeneratedContentResolver.ReferencesAttribute(style.GetValue("content"), attributeName);
}
