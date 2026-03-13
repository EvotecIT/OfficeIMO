using System.Text;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Shared HTML contract for renderer-produced visual blocks such as charts, networks, and dataviews.
/// Hosts can use this API to emit or inspect the same metadata shape as the built-in renderer.
/// </summary>
public static class MarkdownVisualContract {
    /// <summary>Current shared visual contract version.</summary>
    public const string ContractVersion = "v1";

    /// <summary>Config format emitted by the built-in visual renderers.</summary>
    public const string ConfigFormatJson = "json";

    /// <summary>Config encoding emitted by the built-in visual renderers.</summary>
    public const string ConfigEncodingBase64Utf8 = "base64-utf8";

    /// <summary>
    /// Creates a payload descriptor from raw JSON or other renderer-owned source text.
    /// </summary>
    public static MarkdownVisualPayload CreatePayload(string? rawContent) {
        var raw = rawContent ?? string.Empty;
        return new MarkdownVisualPayload(
            raw,
            Convert.ToBase64String(Encoding.UTF8.GetBytes(raw)),
            MarkdownRenderer.ComputeShortHash(raw));
    }

    /// <summary>
    /// Builds a visual host element with the shared <c>data-omd-*</c> metadata contract.
    /// </summary>
    public static string BuildElementHtml(
        string elementName,
        string cssClass,
        string visualKind,
        string language,
        MarkdownVisualPayload payload,
        params KeyValuePair<string, string?>[] extraAttributes) {
        if (string.IsNullOrWhiteSpace(elementName)) {
            throw new ArgumentException("Element name is required.", nameof(elementName));
        }
        if (string.IsNullOrWhiteSpace(cssClass)) {
            throw new ArgumentException("CSS class is required.", nameof(cssClass));
        }
        if (string.IsNullOrWhiteSpace(visualKind)) {
            throw new ArgumentException("Visual kind is required.", nameof(visualKind));
        }
        if (string.IsNullOrWhiteSpace(language)) {
            throw new ArgumentException("Fence language is required.", nameof(language));
        }

        var sb = new StringBuilder();
        sb.Append('<')
          .Append(elementName)
          .Append(" class=\"")
          .Append(HtmlEncode(cssClass))
          .Append('"');
        AppendCommonAttributes(sb, visualKind, language, payload);
        AppendAttributes(sb, extraAttributes);
        sb.Append("></")
          .Append(elementName)
          .Append('>');
        return sb.ToString();
    }

    /// <summary>
    /// Appends the shared visual metadata attributes to an existing HTML element builder.
    /// </summary>
    public static void AppendCommonAttributes(
        StringBuilder sb,
        string visualKind,
        string language,
        MarkdownVisualPayload payload) {
        if (sb == null) {
            throw new ArgumentNullException(nameof(sb));
        }
        if (string.IsNullOrWhiteSpace(visualKind)) {
            throw new ArgumentException("Visual kind is required.", nameof(visualKind));
        }
        if (string.IsNullOrWhiteSpace(language)) {
            throw new ArgumentException("Fence language is required.", nameof(language));
        }

        AppendAttribute(sb, "data-omd-visual-contract", ContractVersion);
        AppendAttribute(sb, "data-omd-visual-kind", visualKind);
        AppendAttribute(sb, "data-omd-fence-language", language);
        AppendAttribute(sb, "data-omd-visual-hash", payload.Hash);
        AppendAttribute(sb, "data-omd-config-format", ConfigFormatJson);
        AppendAttribute(sb, "data-omd-config-encoding", ConfigEncodingBase64Utf8);
        AppendAttribute(sb, "data-omd-config-b64", payload.Base64);
    }

    /// <summary>
    /// Appends an encoded string-valued HTML attribute.
    /// </summary>
    public static void AppendAttribute(StringBuilder sb, string name, string? value) {
        if (sb == null) {
            throw new ArgumentNullException(nameof(sb));
        }

        if (string.IsNullOrWhiteSpace(name) || value == null) {
            return;
        }

        sb.Append(' ')
          .Append(name)
          .Append("=\"")
          .Append(HtmlEncode(value))
          .Append('"');
    }

    /// <summary>
    /// Appends an integer-valued HTML attribute using invariant culture formatting.
    /// </summary>
    public static void AppendAttribute(StringBuilder sb, string name, int value) {
        if (sb == null) {
            throw new ArgumentNullException(nameof(sb));
        }

        if (string.IsNullOrWhiteSpace(name)) {
            return;
        }

        sb.Append(' ')
          .Append(name)
          .Append("=\"")
          .Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append('"');
    }

    /// <summary>
    /// Appends multiple encoded string-valued HTML attributes.
    /// </summary>
    public static void AppendAttributes(StringBuilder sb, params KeyValuePair<string, string?>[] attributes) {
        if (sb == null) {
            throw new ArgumentNullException(nameof(sb));
        }

        if (attributes == null || attributes.Length == 0) {
            return;
        }

        for (int i = 0; i < attributes.Length; i++) {
            var attribute = attributes[i];
            AppendAttribute(sb, attribute.Key, attribute.Value);
        }
    }

    private static string HtmlEncode(string value) {
        return System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
    }
}

/// <summary>
/// Encoded/raw payload data used by the shared renderer visual contract.
/// </summary>
public sealed class MarkdownVisualPayload {
    /// <summary>
    /// Creates a visual payload descriptor.
    /// </summary>
    public MarkdownVisualPayload(string rawContent, string base64, string hash) {
        RawContent = rawContent ?? string.Empty;
        Base64 = base64 ?? string.Empty;
        Hash = hash ?? string.Empty;
    }

    /// <summary>Original raw payload text.</summary>
    public string RawContent { get; }

    /// <summary>UTF-8 base64-encoded payload text.</summary>
    public string Base64 { get; }

    /// <summary>Short stable hash of the raw payload text.</summary>
    public string Hash { get; }
}
