using System.Text;
using System.Text.Json;
using System.Linq;
using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Shared HTML contract for renderer-produced visual blocks such as charts, networks, and dataviews.
/// Hosts can use this API to emit or inspect the same metadata shape as the built-in renderer.
/// </summary>
public static class MarkdownVisualContract {
    /// <summary>Current shared visual contract version.</summary>
    public const string ContractVersion = MarkdownVisualElementContract.ContractVersion;

    /// <summary>Config format emitted by the built-in visual renderers.</summary>
    public const string ConfigFormatJson = MarkdownVisualElementContract.ConfigFormatJson;

    /// <summary>Config encoding emitted by the built-in visual renderers.</summary>
    public const string ConfigEncodingBase64Utf8 = MarkdownVisualElementContract.ConfigEncodingBase64Utf8;

    /// <summary>
    /// Creates a payload descriptor from raw JSON or other renderer-owned source text.
    /// </summary>
    public static MarkdownVisualPayload CreatePayload(string? rawContent) {
        var raw = rawContent ?? string.Empty;
        return new MarkdownVisualPayload(
            raw,
            Convert.ToBase64String(Encoding.UTF8.GetBytes(raw)),
            ComputePayloadHash(raw));
    }

    /// <summary>
    /// Computes a stable cross-platform payload hash for renderer-owned visual content.
    /// JSON payloads are canonicalized before hashing so whitespace and object-property order do not affect the hash.
    /// </summary>
    public static string ComputePayloadHash(string? rawContent) {
        return MarkdownRenderer.ComputeShortHash(CanonicalizeForHashing(rawContent));
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

        AppendAttribute(sb, MarkdownVisualElementContract.AttributeVisualContract, ContractVersion);
        AppendAttribute(sb, MarkdownVisualElementContract.AttributeVisualKind, visualKind);
        AppendAttribute(sb, MarkdownVisualElementContract.AttributeFenceLanguage, language);
        AppendAttribute(sb, MarkdownVisualElementContract.AttributeVisualHash, payload.Hash);
        AppendAttribute(sb, MarkdownVisualElementContract.AttributeConfigFormat, ConfigFormatJson);
        AppendAttribute(sb, MarkdownVisualElementContract.AttributeConfigEncoding, ConfigEncodingBase64Utf8);
        AppendAttribute(sb, MarkdownVisualElementContract.AttributeConfigBase64, payload.Base64);
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

    private static string CanonicalizeForHashing(string? rawContent) {
        var raw = NormalizeLineEndings(rawContent ?? string.Empty);
        if (raw.Length == 0) {
            return string.Empty;
        }

        try {
            using var document = JsonDocument.Parse(raw);
            return CanonicalizeJson(document.RootElement);
        } catch (JsonException) {
            return raw;
        }
    }

    private static string CanonicalizeJson(JsonElement element) {
        return element.ValueKind switch {
            JsonValueKind.Object => CanonicalizeObject(element),
            JsonValueKind.Array => CanonicalizeArray(element),
            JsonValueKind.String => JsonSerializer.Serialize(element.GetString() ?? string.Empty),
            JsonValueKind.Number => element.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => "null",
            _ => element.GetRawText()
        };
    }

    private static string CanonicalizeObject(JsonElement element) {
        var properties = element.EnumerateObject()
            .OrderBy(static property => property.Name, StringComparer.Ordinal)
            .ToArray();
        if (properties.Length == 0) {
            return "{}";
        }

        var sb = new StringBuilder();
        sb.Append('{');
        for (int i = 0; i < properties.Length; i++) {
            if (i > 0) {
                sb.Append(',');
            }

            var property = properties[i];
            sb.Append(JsonSerializer.Serialize(property.Name))
              .Append(':')
              .Append(CanonicalizeJson(property.Value));
        }
        sb.Append('}');
        return sb.ToString();
    }

    private static string CanonicalizeArray(JsonElement element) {
        var sb = new StringBuilder();
        sb.Append('[');
        bool first = true;
        foreach (var item in element.EnumerateArray()) {
            if (!first) {
                sb.Append(',');
            }

            sb.Append(CanonicalizeJson(item));
            first = false;
        }
        sb.Append(']');
        return sb.ToString();
    }

    private static string NormalizeLineEndings(string value) {
        return value
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
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

    /// <summary>Short stable hash of the canonicalized payload text.</summary>
    public string Hash { get; }
}
