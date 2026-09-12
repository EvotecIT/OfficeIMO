using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlImageSourceResolver {
    /// <summary>Resolves the preferred source from an owned image element under the supplied URL policy.</summary>
    public static string ResolveImageSource(Dom.HtmlElement? element, Uri? baseUri, HtmlUrlPolicy? policy, bool allowParentPictureFallback = true) =>
        ResolveImageSource((IElement)NativeDomBridge.GetNativeOrNull(element)!, baseUri, policy, allowParentPictureFallback);
    /// <summary>Resolves ordered image candidates, including permitted picture sources.</summary>
    public static IReadOnlyList<string> ResolveImageSourceCandidates(Dom.HtmlElement? element, Uri? baseUri, HtmlUrlPolicy? policy, bool allowParentPictureFallback = true) =>
        ResolveImageSourceCandidates((IElement)NativeDomBridge.GetNativeOrNull(element)!, baseUri, policy, allowParentPictureFallback);
    /// <summary>Resolves image candidates under an explicit candidate limit.</summary>
    public static IReadOnlyList<string> ResolveImageSourceCandidates(Dom.HtmlElement? element, Uri? baseUri, HtmlUrlPolicy? policy, bool allowParentPictureFallback, int? maxResponsiveCandidates) =>
        ResolveImageSourceCandidates((IElement)NativeDomBridge.GetNativeOrNull(element)!, baseUri, policy, allowParentPictureFallback, maxResponsiveCandidates);
    /// <summary>Resolves the preferred source in an owned picture element.</summary>
    public static string ResolvePictureSource(Dom.HtmlElement? pictureElement, Uri? baseUri, HtmlUrlPolicy? policy) =>
        ResolvePictureSource((IElement)NativeDomBridge.GetNativeOrNull(pictureElement)!, baseUri, policy);
    /// <summary>Resolves ordered picture candidates.</summary>
    public static IReadOnlyList<string> ResolvePictureSourceCandidates(Dom.HtmlElement? pictureElement, Uri? baseUri, HtmlUrlPolicy? policy) =>
        ResolvePictureSourceCandidates((IElement)NativeDomBridge.GetNativeOrNull(pictureElement)!, baseUri, policy);
    /// <summary>Resolves picture candidates under an explicit candidate limit.</summary>
    public static IReadOnlyList<string> ResolvePictureSourceCandidates(Dom.HtmlElement? pictureElement, Uri? baseUri, HtmlUrlPolicy? policy, int? maxCandidates) =>
        ResolvePictureSourceCandidates((IElement)NativeDomBridge.GetNativeOrNull(pictureElement)!, baseUri, policy, maxCandidates);
    /// <summary>Resolves a URL from selected srcset attributes under the URL policy.</summary>
    public static string ResolveUrlFromSrcSetAttributes(Dom.HtmlElement? element, Uri? baseUri, HtmlUrlPolicy? policy, params string[] attributeNames) =>
        ResolveUrlFromSrcSetAttributes((IElement)NativeDomBridge.GetNativeOrNull(element)!, baseUri, policy, attributeNames);
    /// <summary>Normalizes selected srcset attributes under the URL policy.</summary>
    public static string ResolveNormalizedSrcSetAttributes(Dom.HtmlElement? element, Uri? baseUri, HtmlUrlPolicy? policy, params string[] attributeNames) =>
        ResolveNormalizedSrcSetAttributes((IElement)NativeDomBridge.GetNativeOrNull(element)!, baseUri, policy, attributeNames);
    /// <summary>Resolves the first accepted URL from selected attributes.</summary>
    public static string ResolveUrlAttributes(Dom.HtmlElement? element, Uri? baseUri, HtmlUrlPolicy? policy, params string[] attributeNames) =>
        ResolveUrlAttributes((IElement)NativeDomBridge.GetNativeOrNull(element)!, baseUri, policy, attributeNames);
}
