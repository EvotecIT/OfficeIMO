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
        var evaluation = Evaluate(rawUrl, policy, allowEmptyFragment);
        if (!evaluation.IsAllowed) {
            return string.Empty;
        }

        string candidate = evaluation.NormalizedUrl;
        if (!evaluation.AllowBaseUriResolution || baseUri == null) {
            return candidate;
        }

        if (!Uri.TryCreate(baseUri, candidate, out var resolved)) {
            return candidate;
        }

        return IsAllowedResolvedUri(resolved, policy ?? HtmlUrlPolicy.CreateOfficeIMOProfile())
            ? resolved.AbsoluteUri
            : string.Empty;
    }

    private static HtmlUrlEvaluation Evaluate(string? rawUrl, HtmlUrlPolicy? policy, bool allowEmptyFragment) {
        if (string.IsNullOrWhiteSpace(rawUrl)) {
            return HtmlUrlEvaluation.Rejected;
        }

        var effectivePolicy = policy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();
        string candidate = rawUrl!.Trim();
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

    private static bool LooksLikeWindowsDrivePath(string value) {
        return value.Length >= 3
               && char.IsLetter(value[0])
               && value[1] == ':'
               && (value[2] == '\\' || value[2] == '/');
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
