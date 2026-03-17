using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Neutral metadata contract for renderer-produced visual host elements such as charts, networks, and data views.
/// </summary>
public static class MarkdownVisualElementContract {
    /// <summary>Current shared visual contract version.</summary>
    public const string ContractVersion = "v1";

    /// <summary>Shared config format emitted by built-in visual renderers.</summary>
    public const string ConfigFormatJson = "json";

    /// <summary>Shared config encoding emitted by built-in visual renderers.</summary>
    public const string ConfigEncodingBase64Utf8 = "base64-utf8";

    /// <summary>Attribute name storing the contract version.</summary>
    public const string AttributeVisualContract = "data-omd-visual-contract";

    /// <summary>Attribute name storing the visual semantic kind.</summary>
    public const string AttributeVisualKind = "data-omd-visual-kind";

    /// <summary>Attribute name storing the source fence language.</summary>
    public const string AttributeFenceLanguage = "data-omd-fence-language";

    /// <summary>Optional attribute name storing the original normalized fence metadata tail.</summary>
    public const string AttributeFenceInfo = "data-omd-fence-info";

    /// <summary>Optional attribute name storing the original fence element id.</summary>
    public const string AttributeFenceId = "data-omd-fence-id";

    /// <summary>Optional attribute name storing the original fence CSS classes.</summary>
    public const string AttributeFenceClasses = "data-omd-fence-classes";

    /// <summary>Optional attribute name storing a human-friendly visual title.</summary>
    public const string AttributeVisualTitle = "data-omd-visual-title";

    /// <summary>Attribute name storing the visual payload hash.</summary>
    public const string AttributeVisualHash = "data-omd-visual-hash";

    /// <summary>Attribute name storing the payload format.</summary>
    public const string AttributeConfigFormat = "data-omd-config-format";

    /// <summary>Attribute name storing the payload encoding.</summary>
    public const string AttributeConfigEncoding = "data-omd-config-encoding";

    /// <summary>Attribute name storing the base64 payload content.</summary>
    public const string AttributeConfigBase64 = "data-omd-config-b64";

    /// <summary>
    /// Parses a visual host attribute set into a typed descriptor.
    /// </summary>
    public static bool TryParse(IEnumerable<KeyValuePair<string, string?>> attributes, out MarkdownVisualElement? element) {
        element = null;
        if (attributes == null) {
            return false;
        }

        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var attribute in attributes) {
            if (string.IsNullOrWhiteSpace(attribute.Key) || attribute.Value == null) {
                continue;
            }

            values[attribute.Key] = attribute.Value;
        }

        if (!TryGetRequired(values, AttributeVisualContract, out var contractVersion)
            || !TryGetRequired(values, AttributeVisualKind, out var visualKind)
            || !TryGetRequired(values, AttributeFenceLanguage, out var fenceLanguage)
            || !TryGetRequired(values, AttributeVisualHash, out var visualHash)
            || !TryGetRequired(values, AttributeConfigFormat, out var configFormat)
            || !TryGetRequired(values, AttributeConfigEncoding, out var configEncoding)
            || !TryGetRequired(values, AttributeConfigBase64, out var configBase64)) {
            return false;
        }

        element = new MarkdownVisualElement(
            contractVersion,
            visualKind,
            fenceLanguage,
            visualHash,
            configFormat,
            configEncoding,
            configBase64,
            new Dictionary<string, string>(values, StringComparer.OrdinalIgnoreCase));
        return true;
    }

    /// <summary>
    /// Decodes a payload string when the descriptor uses the standard JSON/base64 contract.
    /// </summary>
    public static string? TryDecodePayload(MarkdownVisualElement? element) {
        if (element == null) {
            return null;
        }

        if (!string.Equals(element.ConfigEncoding, ConfigEncodingBase64Utf8, StringComparison.OrdinalIgnoreCase)) {
            return null;
        }

        try {
            var bytes = Convert.FromBase64String(element.ConfigBase64);
            return Encoding.UTF8.GetString(bytes);
        } catch (FormatException) {
            return null;
        } catch (ArgumentException) {
            return null;
        }
    }

    private static bool TryGetRequired(IReadOnlyDictionary<string, string> values, string key, out string value) {
        if (values.TryGetValue(key, out var candidate) && !string.IsNullOrWhiteSpace(candidate)) {
            value = candidate;
            return true;
        }

        value = string.Empty;
        return false;
    }
}

/// <summary>
/// Typed descriptor for a renderer-produced visual host element.
/// </summary>
public sealed class MarkdownVisualElement {
    internal MarkdownVisualElement(
        string contractVersion,
        string visualKind,
        string fenceLanguage,
        string visualHash,
        string configFormat,
        string configEncoding,
        string configBase64,
        IReadOnlyDictionary<string, string> attributes) {
        ContractVersion = contractVersion ?? string.Empty;
        VisualKind = visualKind ?? string.Empty;
        FenceLanguage = fenceLanguage ?? string.Empty;
        VisualHash = visualHash ?? string.Empty;
        ConfigFormat = configFormat ?? string.Empty;
        ConfigEncoding = configEncoding ?? string.Empty;
        ConfigBase64 = configBase64 ?? string.Empty;
        Attributes = attributes ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Shared contract version.</summary>
    public string ContractVersion { get; }

    /// <summary>Visual semantic kind such as <c>chart</c> or <c>network</c>.</summary>
    public string VisualKind { get; }

    /// <summary>Original source fence language.</summary>
    public string FenceLanguage { get; }

    /// <summary>Original normalized fence metadata tail after the primary language token.</summary>
    public string? FenceAdditionalInfo => TryGetAttribute(MarkdownVisualElementContract.AttributeFenceInfo, out var value) ? value : null;

    /// <summary>Original source fence element id when the fence metadata provided one.</summary>
    public string? FenceElementId => TryGetAttribute(MarkdownVisualElementContract.AttributeFenceId, out var value) ? value : null;

    /// <summary>Original source fence CSS classes when the fence metadata provided them.</summary>
    public IReadOnlyList<string> FenceClasses {
        get {
            if (!TryGetAttribute(MarkdownVisualElementContract.AttributeFenceClasses, out var value) || string.IsNullOrWhiteSpace(value)) {
                return Array.Empty<string>();
            }

            var classes = new List<string>();
            foreach (var className in value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
                bool exists = false;
                for (int i = 0; i < classes.Count; i++) {
                    if (string.Equals(classes[i], className, StringComparison.OrdinalIgnoreCase)) {
                        exists = true;
                        break;
                    }
                }

                if (exists) {
                    continue;
                }

                classes.Add(className);
            }

            return classes;
        }
    }

    /// <summary>Optional human-friendly visual title when the source fence metadata provides one.</summary>
    public string? VisualTitle => TryGetAttribute(MarkdownVisualElementContract.AttributeVisualTitle, out var value) ? value : null;

    /// <summary>Stable payload hash.</summary>
    public string VisualHash { get; }

    /// <summary>Declared payload format.</summary>
    public string ConfigFormat { get; }

    /// <summary>Declared payload encoding.</summary>
    public string ConfigEncoding { get; }

    /// <summary>Encoded payload content.</summary>
    public string ConfigBase64 { get; }

    /// <summary>All parsed attributes from the host element.</summary>
    public IReadOnlyDictionary<string, string> Attributes { get; }

    /// <summary>
    /// Attempts to read an attribute value from the parsed host element.
    /// </summary>
    public bool TryGetAttribute(string name, out string value) {
        if (string.IsNullOrWhiteSpace(name)) {
            value = string.Empty;
            return false;
        }

        return Attributes.TryGetValue(name, out value!);
    }

    /// <summary>
    /// Decodes the payload content when the descriptor uses the standard JSON/base64 contract.
    /// </summary>
    public string? TryDecodePayload() {
        return MarkdownVisualElementContract.TryDecodePayload(this);
    }
}
