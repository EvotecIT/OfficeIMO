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
        foreach (string candidate in ResolveImageSourceCandidates(element, baseUri, policy, allowParentPictureFallback)) {
            return candidate;
        }

        return string.Empty;
    }

    /// <summary>
    /// Resolves all image source candidates from an element in browser-like preference order.
    /// </summary>
    public static IReadOnlyList<string> ResolveImageSourceCandidates(IElement element, Uri? baseUri, HtmlUrlPolicy? policy, bool allowParentPictureFallback = true) {
        return ResolveImageSourceCandidates(element, baseUri, policy, allowParentPictureFallback, null);
    }

    /// <summary>
    /// Resolves image source candidates from an element in browser-like preference order, limiting responsive candidates from picture/source-set inputs.
    /// </summary>
    public static IReadOnlyList<string> ResolveImageSourceCandidates(IElement element, Uri? baseUri, HtmlUrlPolicy? policy, bool allowParentPictureFallback, int? maxResponsiveCandidates) {
        var candidates = new CandidateAccumulator();
        if (element == null) {
            return candidates.Items;
        }

        int responsiveCandidateCount = 0;
        if (allowParentPictureFallback
            && element.ParentElement != null
            && element.ParentElement.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)) {
            AddPictureSourceCandidates(candidates, element.ParentElement, baseUri, policy, maxResponsiveCandidates, ref responsiveCandidateCount);
        }

        AddResolvedUrlAttributes(candidates, element, baseUri, policy, LazySourceAttributes);
        AddResolvedSrcSetAttributes(candidates, element, baseUri, policy, maxResponsiveCandidates, ref responsiveCandidateCount, SrcSetAttributes);
        AddResolvedUrlAttributes(candidates, element, baseUri, policy, SourceAttributes);

        return candidates.Items;
    }

    /// <summary>
    /// Resolves the candidates selected by the active render media environment.
    /// </summary>
    internal static IReadOnlyList<string> ResolveImageSourceCandidatesForRendering(IElement element, Uri? baseUri, HtmlUrlPolicy? policy, HtmlRenderOptions options) {
        var candidates = new CandidateAccumulator();
        if (element == null) return candidates.Items;

        bool selectedPictureSource = false;
        IElement? picture = element.ParentElement;
        if (picture != null && picture.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)) {
            double mediaWidth = options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth;
            double mediaHeight = options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D;
            foreach (IElement child in picture.Children) {
                if (ReferenceEquals(child, element)) break;
                if (!child.TagName.Equals("SOURCE", StringComparison.OrdinalIgnoreCase)
                    || !HtmlComputedStyleEngine.IsApplicableMedia(child.GetAttribute("media") ?? string.Empty, options.MediaContext, mediaWidth, mediaHeight)
                    || !HtmlPictureSourceSupport.IsSupportedConversionContentType(child.GetAttribute("type"))) {
                    continue;
                }

                int candidateCount = 0;
                int countBeforeSource = candidates.Items.Count;
                AddResolvedSrcSetAttributes(candidates, child, baseUri, policy, options.ResponsiveImageCandidateLimit, ref candidateCount, SrcSetAttributes);
                AddResolvedUrlAttributes(candidates, child, baseUri, policy, options.ResponsiveImageCandidateLimit, ref candidateCount, PictureSourceAttributes);
                if (candidates.Items.Count > countBeforeSource) {
                    selectedPictureSource = true;
                    break;
                }
            }
        }

        if (!selectedPictureSource) {
            AddResolvedUrlAttributes(candidates, element, baseUri, policy, LazySourceAttributes);
            int responsiveCandidateCount = 0;
            AddResolvedSrcSetAttributes(candidates, element, baseUri, policy, options.ResponsiveImageCandidateLimit, ref responsiveCandidateCount, SrcSetAttributes);
            AddResolvedUrlAttributes(candidates, element, baseUri, policy, SourceAttributes);
        }

        return candidates.Items;
    }

    /// <summary>
    /// Resolves the preferred source from a <c>picture</c> element.
    /// </summary>
    public static string ResolvePictureSource(IElement pictureElement, Uri? baseUri, HtmlUrlPolicy? policy) {
        foreach (string candidate in ResolvePictureSourceCandidates(pictureElement, baseUri, policy)) {
            return candidate;
        }

        return string.Empty;
    }

    /// <summary>
    /// Resolves all source candidates from a <c>picture</c> element.
    /// </summary>
    public static IReadOnlyList<string> ResolvePictureSourceCandidates(IElement pictureElement, Uri? baseUri, HtmlUrlPolicy? policy) {
        return ResolvePictureSourceCandidates(pictureElement, baseUri, policy, null);
    }

    /// <summary>
    /// Resolves source candidates from a <c>picture</c> element, stopping after the requested number of responsive candidates.
    /// </summary>
    public static IReadOnlyList<string> ResolvePictureSourceCandidates(IElement pictureElement, Uri? baseUri, HtmlUrlPolicy? policy, int? maxCandidates) {
        var candidates = new CandidateAccumulator();
        if (pictureElement == null) {
            return candidates.Items;
        }

        int candidateCount = 0;
        AddPictureSourceCandidates(candidates, pictureElement, baseUri, policy, maxCandidates, ref candidateCount);
        return candidates.Items;
    }

    private static void AddPictureSourceCandidates(CandidateAccumulator candidates, IElement pictureElement, Uri? baseUri, HtmlUrlPolicy? policy, int? maxCandidates, ref int candidateCount) {
        foreach (var child in pictureElement.Children) {
            if (!child.TagName.Equals("SOURCE", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            if (!AddResolvedSrcSetAttributes(candidates, child, baseUri, policy, maxCandidates, ref candidateCount, SrcSetAttributes)
                || !AddResolvedUrlAttributes(candidates, child, baseUri, policy, maxCandidates, ref candidateCount, PictureSourceAttributes)) {
                return;
            }
        }
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
        return ResolveNormalizedSrcSet(rawSrcSet, baseUri, policy, null);
    }

    /// <summary>Resolves and normalizes allowed candidates up to the supplied parsing limit.</summary>
    public static string ResolveNormalizedSrcSet(string? rawSrcSet, Uri? baseUri, HtmlUrlPolicy? policy, int? maxCandidates) {
        if (string.IsNullOrWhiteSpace(rawSrcSet)) {
            return string.Empty;
        }

        var parts = new List<string>();
        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(rawSrcSet, maxCandidates)) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(candidate.Url, baseUri, policy);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                parts.Add(string.IsNullOrWhiteSpace(candidate.Descriptor) ? resolved : resolved + " " + candidate.Descriptor);
            }
        }

        return string.Join(", ", parts);
    }

    /// <summary>
    /// Resolves all allowed candidates from a <c>srcset</c> value.
    /// </summary>
    public static IReadOnlyList<HtmlSrcSetCandidate> ResolveSrcSetCandidates(string? rawSrcSet, Uri? baseUri, HtmlUrlPolicy? policy) {
        return ResolveSrcSetCandidates(rawSrcSet, baseUri, policy, null);
    }

    /// <summary>
    /// Resolves allowed candidates from a <c>srcset</c> value, stopping after the requested number of parsed candidates.
    /// </summary>
    public static IReadOnlyList<HtmlSrcSetCandidate> ResolveSrcSetCandidates(string? rawSrcSet, Uri? baseUri, HtmlUrlPolicy? policy, int? maxCandidates) {
        var candidates = new List<HtmlSrcSetCandidate>();
        if (string.IsNullOrWhiteSpace(rawSrcSet) || IsNonPositiveCandidateLimit(maxCandidates)) {
            return candidates;
        }

        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(rawSrcSet, maxCandidates)) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(candidate.Url, baseUri, policy);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                candidates.Add(new HtmlSrcSetCandidate(resolved, candidate.Descriptor));
            }
        }

        return candidates;
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

    private static void AddResolvedSrcSetAttributes(CandidateAccumulator candidates, IElement element, Uri? baseUri, HtmlUrlPolicy? policy, params string[] attributeNames) {
        int candidateCount = 0;
        AddResolvedSrcSetAttributes(candidates, element, baseUri, policy, null, ref candidateCount, attributeNames);
    }

    private static bool AddResolvedSrcSetAttributes(CandidateAccumulator candidates, IElement element, Uri? baseUri, HtmlUrlPolicy? policy, int? maxCandidates, ref int candidateCount, params string[] attributeNames) {
        if (element == null || attributeNames == null || attributeNames.Length == 0) {
            return true;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            int? remaining = GetRemainingCandidateCount(maxCandidates, candidateCount);
            if (IsNonPositiveCandidateLimit(remaining)) {
                return false;
            }

            foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Enumerate(element.GetAttribute(attributeNames[i]))) {
                string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(candidate.Url, baseUri, policy);
                AddCandidate(candidates, resolved);
                candidateCount++;

                if (IsNonPositiveCandidateLimit(GetRemainingCandidateCount(maxCandidates, candidateCount))) {
                    return false;
                }
            }
        }

        return true;
    }

    private static void AddResolvedUrlAttributes(CandidateAccumulator candidates, IElement element, Uri? baseUri, HtmlUrlPolicy? policy, params string[] attributeNames) {
        int candidateCount = 0;
        AddResolvedUrlAttributes(candidates, element, baseUri, policy, null, ref candidateCount, attributeNames);
    }

    private static bool AddResolvedUrlAttributes(CandidateAccumulator candidates, IElement element, Uri? baseUri, HtmlUrlPolicy? policy, int? maxCandidates, ref int candidateCount, params string[] attributeNames) {
        if (element == null || attributeNames == null || attributeNames.Length == 0) {
            return true;
        }

        for (int i = 0; i < attributeNames.Length; i++) {
            if (IsNonPositiveCandidateLimit(GetRemainingCandidateCount(maxCandidates, candidateCount))) {
                return false;
            }

            string? rawValue = element.GetAttribute(attributeNames[i]);
            if (string.IsNullOrWhiteSpace(rawValue)) {
                continue;
            }

            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(rawValue, baseUri, policy);
            AddCandidate(candidates, resolved);
            if (!string.IsNullOrWhiteSpace(rawValue)) {
                candidateCount++;
            }

            if (IsNonPositiveCandidateLimit(GetRemainingCandidateCount(maxCandidates, candidateCount))) {
                return false;
            }
        }

        return true;
    }

    private static int? GetRemainingCandidateCount(int? maxCandidates, int candidateCount) {
        if (!maxCandidates.HasValue) {
            return null;
        }

        return Math.Max(0, maxCandidates.Value - candidateCount);
    }

    private static bool IsNonPositiveCandidateLimit(int? maxCandidates) {
        return maxCandidates.HasValue && maxCandidates.Value <= 0;
    }

    private static bool AddCandidate(CandidateAccumulator candidates, string? source) {
        return candidates.Add(source);
    }

    private sealed class CandidateAccumulator {
        private readonly List<string> _items = new List<string>();
        private readonly HashSet<string> _seen = new HashSet<string>(StringComparer.Ordinal);

        internal IReadOnlyList<string> Items => _items;

        internal bool Add(string? source) {
            if (string.IsNullOrWhiteSpace(source)) {
                return false;
            }

            string candidate = source!;
            if (_seen.Add(candidate)) {
                _items.Add(candidate);
                return true;
            }

            return false;
        }
    }
}
