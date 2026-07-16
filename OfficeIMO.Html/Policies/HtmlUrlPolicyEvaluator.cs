namespace OfficeIMO.Html;

/// <summary>
/// Evaluates and resolves URL values against a shared <see cref="HtmlUrlPolicy"/>.
/// </summary>
public static class HtmlUrlPolicyEvaluator {
    /// <summary>
    /// Returns <see langword="true"/> when the raw URL value is allowed by the supplied policy.
    /// </summary>
    public static bool IsAllowed(string? rawUrl, HtmlUrlPolicy? policy, bool allowEmptyFragment = true) =>
        Evaluate(rawUrl, policy, allowEmptyFragment).IsAllowed;

    /// <summary>
    /// Resolves a raw URL against an optional base URI and returns an empty string when policy rejects it.
    /// </summary>
    public static string ResolveUrl(string? rawUrl, Uri? baseUri, HtmlUrlPolicy? policy, bool allowEmptyFragment = true) {
        var effectivePolicy = policy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();
        var evaluation = Evaluate(rawUrl, effectivePolicy, allowEmptyFragment);
        if (!evaluation.IsAllowed) {
            return string.Empty;
        }

        string candidate = evaluation.NormalizedUrl;
        if (candidate.StartsWith("//", StringComparison.Ordinal)) {
            return ApplyResolvedUrlTransform(
                ResolveProtocolRelativeUrl(candidate, baseUri, effectivePolicy),
                effectivePolicy,
                allowEmptyFragment);
        }

        if (!evaluation.AllowBaseUriResolution || baseUri == null) {
            return ApplyResolvedUrlTransform(candidate, effectivePolicy, allowEmptyFragment);
        }

        if (!Uri.TryCreate(baseUri, candidate, out var resolved)) {
            return ApplyResolvedUrlTransform(candidate, effectivePolicy, allowEmptyFragment);
        }

        string resolvedValue = IsAllowedResolvedUri(resolved, effectivePolicy)
            ? resolved.AbsoluteUri
            : string.Empty;
        return ApplyResolvedUrlTransform(resolvedValue, effectivePolicy, allowEmptyFragment);
    }

    private static string ApplyResolvedUrlTransform(string value, HtmlUrlPolicy policy, bool allowEmptyFragment) {
        if (value.Length == 0 || policy.ResolvedUrlTransform == null) return value;
        string? transformed = policy.ResolvedUrlTransform(value);
        if (string.IsNullOrWhiteSpace(transformed)) return string.Empty;
        HtmlUrlEvaluation evaluation = Evaluate(transformed, policy, allowEmptyFragment);
        return evaluation.IsAllowed ? evaluation.NormalizedUrl : string.Empty;
    }

    private static string ResolveProtocolRelativeUrl(string candidate, Uri? baseUri, HtmlUrlPolicy policy) {
        string scheme = baseUri != null
                        && (baseUri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                            || baseUri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            ? baseUri.Scheme
            : Uri.UriSchemeHttps;

        if (!Uri.TryCreate(scheme + ":" + candidate, UriKind.Absolute, out var resolved)) {
            return candidate;
        }

        return IsAllowedResolvedUri(resolved, policy)
            ? resolved.AbsoluteUri
            : string.Empty;
    }

    private static HtmlUrlEvaluation Evaluate(string? rawUrl, HtmlUrlPolicy? policy, bool allowEmptyFragment) {
        if (string.IsNullOrWhiteSpace(rawUrl)) {
            return HtmlUrlEvaluation.Rejected;
        }

        var effectivePolicy = policy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();
        string candidate = rawUrl!.Trim();
        if (ContainsUrlControlCharacter(candidate)) {
            return HtmlUrlEvaluation.Rejected;
        }

        if (candidate.StartsWith("#", StringComparison.Ordinal)) {
            return allowEmptyFragment || candidate.Length > 1
                ? HtmlUrlEvaluation.Allowed(candidate, allowBaseUriResolution: false)
                : HtmlUrlEvaluation.Rejected;
        }

        if (candidate.StartsWith("//", StringComparison.Ordinal)) {
            return effectivePolicy.AllowProtocolRelativeUrls
                ? HtmlUrlEvaluation.Allowed(candidate, allowBaseUriResolution: true)
                : HtmlUrlEvaluation.Rejected;
        }

        if (candidate.StartsWith("/", StringComparison.Ordinal)) {
            return HtmlUrlEvaluation.Allowed(candidate, allowBaseUriResolution: true);
        }

        if (LooksLikeWindowsDrivePath(candidate)) {
            return effectivePolicy.DisallowFileUrls
                ? HtmlUrlEvaluation.Rejected
                : HtmlUrlEvaluation.Allowed(candidate, allowBaseUriResolution: true);
        }

        if (LooksLikeUncPath(candidate)) {
            return effectivePolicy.DisallowFileUrls
                ? HtmlUrlEvaluation.Rejected
                : HtmlUrlEvaluation.Allowed(candidate, allowBaseUriResolution: false);
        }

        if (Uri.TryCreate(candidate, UriKind.Absolute, out var absoluteUri)) {
            return IsAllowedResolvedUri(absoluteUri, effectivePolicy)
                ? HtmlUrlEvaluation.Allowed(candidate, allowBaseUriResolution: false)
                : HtmlUrlEvaluation.Rejected;
        }

        return HtmlUrlEvaluation.Allowed(candidate, allowBaseUriResolution: true);
    }

    private static bool IsAllowedResolvedUri(Uri uri, HtmlUrlPolicy policy) {
        string scheme = uri.Scheme ?? string.Empty;
        if (scheme.Length == 0) {
            return true;
        }

        if (policy.DisallowScriptUrls
            && (scheme.Equals("javascript", StringComparison.OrdinalIgnoreCase)
                || scheme.Equals("vbscript", StringComparison.OrdinalIgnoreCase))) {
            return false;
        }

        if (policy.DisallowFileUrls && scheme.Equals(Uri.UriSchemeFile, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        if (!policy.AllowMailtoUrls && scheme.Equals(Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        if (!policy.AllowDataUrls && scheme.Equals("data", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return !policy.RestrictUrlSchemes || policy.AllowedUrlSchemes.Contains(scheme);
    }

    private static bool ContainsUrlControlCharacter(string value) {
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (c <= '\u001F' || c == '\u007F') {
                return true;
            }
        }

        return false;
    }

    private static bool LooksLikeWindowsDrivePath(string value) {
        return value.Length >= 2
               && char.IsLetter(value[0])
               && value[1] == ':';
    }

    private static bool LooksLikeUncPath(string value) {
        return value.Length >= 2
               && value[0] == '\\'
               && value[1] == '\\';
    }

    private readonly struct HtmlUrlEvaluation {
        private HtmlUrlEvaluation(bool isAllowed, string normalizedUrl, bool allowBaseUriResolution) {
            IsAllowed = isAllowed;
            NormalizedUrl = normalizedUrl;
            AllowBaseUriResolution = allowBaseUriResolution;
        }

        public static HtmlUrlEvaluation Rejected => new HtmlUrlEvaluation(false, string.Empty, false);

        public static HtmlUrlEvaluation Allowed(string normalizedUrl, bool allowBaseUriResolution) =>
            new HtmlUrlEvaluation(true, normalizedUrl, allowBaseUriResolution);

        public bool IsAllowed { get; }
        public string NormalizedUrl { get; }
        public bool AllowBaseUriResolution { get; }
    }
}
