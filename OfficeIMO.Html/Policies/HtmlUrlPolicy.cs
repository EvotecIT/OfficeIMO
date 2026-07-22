namespace OfficeIMO.Html;

/// <summary>
/// Shared URL policy used by OfficeIMO HTML ingestion adapters before links or resource references are materialized.
/// </summary>
public sealed class HtmlUrlPolicy {
    /// <summary>
    /// Creates the default OfficeIMO URL policy for trusted compatibility-oriented HTML ingestion.
    /// </summary>
    public static HtmlUrlPolicy CreateOfficeIMOProfile() => new HtmlUrlPolicy();

    /// <summary>
    /// Creates a hyperlink policy that rejects script, file, and data URLs while allowing common web, mail, and telephone links.
    /// </summary>
    public static HtmlUrlPolicy CreateHyperlinkProfile() => new HtmlUrlPolicy {
        DisallowFileUrls = true,
        AllowDataUrls = false,
        RestrictUrlSchemes = true,
        AllowedUrlSchemes = { "tel" }
    };

    /// <summary>
    /// Creates a restrictive web-only policy that accepts only HTTP, HTTPS, mail, and telephone links.
    /// </summary>
    public static HtmlUrlPolicy CreateWebOnlyProfile() => new HtmlUrlPolicy {
        DisallowFileUrls = true,
        AllowDataUrls = false,
        RestrictUrlSchemes = true,
        AllowedUrlSchemes = { "tel" }
    };

    /// <summary>
    /// Creates an offline resource policy that accepts bounded embedded data while rejecting file and network schemes.
    /// </summary>
    public static HtmlUrlPolicy CreateEmbeddedResourceProfile() {
        var policy = new HtmlUrlPolicy {
            DisallowFileUrls = true,
            AllowMailtoUrls = false,
            AllowDataUrls = true,
            AllowProtocolRelativeUrls = false,
            RestrictUrlSchemes = true
        };
        policy.AllowedUrlSchemes.Clear();
        policy.AllowedUrlSchemes.Add("data");
        return policy;
    }

    /// <summary>
    /// When true, script-like URL schemes such as <c>javascript:</c> and <c>vbscript:</c> are rejected.
    /// </summary>
    public bool DisallowScriptUrls { get; set; } = true;

    /// <summary>
    /// When true, <c>file:</c> URLs and Windows drive-path URL targets are rejected.
    /// </summary>
    public bool DisallowFileUrls { get; set; }

    /// <summary>
    /// When false, <c>mailto:</c> URLs are rejected.
    /// </summary>
    public bool AllowMailtoUrls { get; set; } = true;

    /// <summary>
    /// When false, <c>data:</c> URLs are rejected.
    /// </summary>
    public bool AllowDataUrls { get; set; } = true;

    /// <summary>
    /// When false, protocol-relative URLs such as <c>//example.com/image.png</c> are rejected unless callers resolve them first.
    /// </summary>
    public bool AllowProtocolRelativeUrls { get; set; } = true;

    /// <summary>
    /// When true, absolute URLs must use a scheme listed in <see cref="AllowedUrlSchemes"/>.
    /// Relative URLs and fragments are still allowed by the evaluator.
    /// </summary>
    public bool RestrictUrlSchemes { get; set; }

    /// <summary>
    /// Optional trusted transform applied after a URL passes policy checks and base-URI resolution.
    /// Return <see langword="null"/> or an empty string to reject the resolved URL. Transformed
    /// values are checked against this policy again before use.
    /// </summary>
    public Func<string, string?>? ResolvedUrlTransform { get; set; }

    /// <summary>
    /// URL schemes allowed when <see cref="RestrictUrlSchemes"/> is enabled.
    /// </summary>
    public HashSet<string> AllowedUrlSchemes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        Uri.UriSchemeHttp,
        Uri.UriSchemeHttps,
        Uri.UriSchemeMailto
    };

    /// <summary>
    /// Creates a copy of the current policy instance.
    /// </summary>
    public HtmlUrlPolicy Clone() {
        var clone = new HtmlUrlPolicy {
            DisallowScriptUrls = DisallowScriptUrls,
            DisallowFileUrls = DisallowFileUrls,
            AllowMailtoUrls = AllowMailtoUrls,
            AllowDataUrls = AllowDataUrls,
            AllowProtocolRelativeUrls = AllowProtocolRelativeUrls,
            RestrictUrlSchemes = RestrictUrlSchemes,
            ResolvedUrlTransform = ResolvedUrlTransform
        };

        clone.AllowedUrlSchemes.Clear();
        foreach (string scheme in AllowedUrlSchemes) {
            if (!string.IsNullOrWhiteSpace(scheme)) {
                clone.AllowedUrlSchemes.Add(scheme.Trim());
            }
        }

        return clone;
    }

    /// <summary>Creates the restrictive intersection of two independently supplied policy boundaries.</summary>
    internal static HtmlUrlPolicy Intersect(HtmlUrlPolicy? first, HtmlUrlPolicy? second) {
        HtmlUrlPolicy left = first ?? CreateOfficeIMOProfile();
        HtmlUrlPolicy right = second ?? CreateOfficeIMOProfile();
        var result = new HtmlUrlPolicy {
            DisallowScriptUrls = left.DisallowScriptUrls || right.DisallowScriptUrls,
            DisallowFileUrls = left.DisallowFileUrls || right.DisallowFileUrls,
            AllowMailtoUrls = left.AllowMailtoUrls && right.AllowMailtoUrls,
            AllowDataUrls = left.AllowDataUrls && right.AllowDataUrls,
            AllowProtocolRelativeUrls = left.AllowProtocolRelativeUrls && right.AllowProtocolRelativeUrls,
            RestrictUrlSchemes = left.RestrictUrlSchemes || right.RestrictUrlSchemes,
            ResolvedUrlTransform = ComposeTransforms(left.ResolvedUrlTransform, right.ResolvedUrlTransform)
        };

        result.AllowedUrlSchemes.Clear();
        IEnumerable<string> schemes = left.RestrictUrlSchemes && right.RestrictUrlSchemes
            ? left.AllowedUrlSchemes.Intersect(right.AllowedUrlSchemes, StringComparer.OrdinalIgnoreCase)
            : left.RestrictUrlSchemes ? left.AllowedUrlSchemes
            : right.RestrictUrlSchemes ? right.AllowedUrlSchemes
            : left.AllowedUrlSchemes.Union(right.AllowedUrlSchemes, StringComparer.OrdinalIgnoreCase);
        foreach (string scheme in schemes) result.AllowedUrlSchemes.Add(scheme);
        return result;
    }

    private static Func<string, string?>? ComposeTransforms(
        Func<string, string?>? first,
        Func<string, string?>? second) {
        if (first == null) return second;
        if (second == null) return first;
        return value => {
            string? transformed = first(value);
            return string.IsNullOrWhiteSpace(transformed) ? null : second(transformed!);
        };
    }
}
