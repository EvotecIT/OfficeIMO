using AngleSharp.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Resolves image sources from common HTML image, lazy-loading, and picture/source patterns.
/// </summary>
public static class HtmlImageSourceResolver {
    private static readonly string[] LazySourceAttributes = { "data-src", "data-original", "data-original-src", "data-lazy-src" };
    private static readonly string[] SourceAttributes = { "src" };
    private static readonly string[] SrcSetAttributes = { "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset" };
    private static readonly string[] PictureSourceAttributes = { "src", "data-src", "data-original-src", "data-lazy-src" };

    /// <summary>
    /// Resolves the best image source from an element, including lazy-loading attributes, source sets, and parent picture fallbacks.
    /// </summary>
    public static string ResolveImageSource(IElement element, Uri? baseUri, HtmlUrlPolicy? policy, bool allowParentPictureFallback = true) {
        if (element == null) {
            return string.Empty;
        }

        string resolved = ResolveUrlAttributes(element, baseUri, policy, LazySourceAttributes);
        if (!string.IsNullOrWhiteSpace(resolved)) {
            return resolved;
        }

        resolved = ResolveUrlAttributes(element, baseUri, policy, SourceAttributes);
        if (!string.IsNullOrWhiteSpace(resolved)) {
            return resolved;
        }

        resolved = ResolveUrlFromSrcSetAttributes(element, baseUri, policy, SrcSetAttributes);
        if (!string.IsNullOrWhiteSpace(resolved)) {
            return resolved;
        }

        return allowParentPictureFallback
            && element.ParentElement != null
            && element.ParentElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)
            ? ResolvePictureSource(element.ParentElement, baseUri, policy)
            : string.Empty;
    }

    /// <summary>
    /// Resolves the preferred source from a <c>picture</c> element.
    /// </summary>
    public static string ResolvePictureSource(IElement pictureElement, Uri? baseUri, HtmlUrlPolicy? policy) {
        if (pictureElement == null) {
            return string.Empty;
        }

        foreach (var child in pictureElement.Children) {
            if (!child.TagName.Equals("SOURCE", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string resolved = ResolveUrlFromSrcSetAttributes(child, baseUri, policy, SrcSetAttributes);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }

            resolved = ResolveUrlAttributes(child, baseUri, policy, PictureSourceAttributes);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        return string.Empty;
    }

    /// <summary>
    /// Resolves the first allowed candidate from a <c>srcset</c> value.
    /// </summary>
    public static string ResolveUrlFromSrcSet(string? rawSrcSet, Uri? baseUri, HtmlUrlPolicy? policy) {
        return ResolveFirstSrcSetCandidate(rawSrcSet, baseUri, policy).Url;
    }

    /// <summary>
    /// Resolves and normalizes all allowed candidates from a <c>srcset</c> value.
    /// </summary>
    public static string ResolveNormalizedSrcSet(string? rawSrcSet, Uri? baseUri, HtmlUrlPolicy? policy) {
        if (string.IsNullOrWhiteSpace(rawSrcSet)) {
            return string.Empty;
        }

        var parts = new List<string>();
        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(rawSrcSet)) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(candidate.Url, baseUri, policy);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                parts.Add(string.IsNullOrWhiteSpace(candidate.Descriptor) ? resolved : resolved + " " + candidate.Descriptor);
            }
        }

        return string.Join(", ", parts);
    }

    /// <summary>
    /// Resolves the first allowed candidate from each supplied source-set attribute in order.
    /// </summary>
    public static string ResolveUrlFromSrcSetAttributes(IElement element, Uri? baseUri, HtmlUrlPolicy? policy, params string[] attributeNames) {
        if (element == null || attributeNames == null || attributeNames.Length == 0) {
            return string.Empty;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            string resolved = ResolveUrlFromSrcSet(element.GetAttribute(attributeNames[i]), baseUri, policy);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        return string.Empty;
    }

    /// <summary>
    /// Resolves and normalizes the first non-empty source-set attribute.
    /// </summary>
    public static string ResolveNormalizedSrcSetAttributes(IElement element, Uri? baseUri, HtmlUrlPolicy? policy, params string[] attributeNames) {
        if (element == null || attributeNames == null || attributeNames.Length == 0) {
            return string.Empty;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            string resolved = ResolveNormalizedSrcSet(element.GetAttribute(attributeNames[i]), baseUri, policy);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        return string.Empty;
    }

    /// <summary>
    /// Resolves the first allowed URL attribute from an element.
    /// </summary>
    public static string ResolveUrlAttributes(IElement element, Uri? baseUri, HtmlUrlPolicy? policy, params string[] attributeNames) {
        if (element == null || attributeNames == null || attributeNames.Length == 0) {
            return string.Empty;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(element.GetAttribute(attributeNames[i]), baseUri, policy);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return resolved;
            }
        }

        return string.Empty;
    }

    /// <summary>
    /// Resolves the first allowed source-set candidate and preserves its descriptor.
    /// </summary>
    public static HtmlSrcSetCandidate ResolveFirstSrcSetCandidate(string? rawSrcSet, Uri? baseUri, HtmlUrlPolicy? policy) {
        if (string.IsNullOrWhiteSpace(rawSrcSet)) {
            return new HtmlSrcSetCandidate(string.Empty, string.Empty);
        }

        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(rawSrcSet)) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(candidate.Url, baseUri, policy);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return new HtmlSrcSetCandidate(resolved, candidate.Descriptor);
            }
        }

        return new HtmlSrcSetCandidate(string.Empty, string.Empty);
    }
}
