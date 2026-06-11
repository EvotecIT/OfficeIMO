namespace OfficeIMO.Markdown;

/// <summary>
/// Dependency-free typed metadata for a semantic visual fenced-block payload.
/// </summary>
public sealed class MarkdownNativeVisualPayload {
    private MarkdownNativeVisualPayload(
        MarkdownNativeVisualPayloadFormat format,
        string declaredSemanticKind,
        string? detectedSemanticKind,
        string? jsonType,
        IReadOnlyDictionary<string, string> signals) {
        Format = format;
        DeclaredSemanticKind = declaredSemanticKind ?? string.Empty;
        DetectedSemanticKind = detectedSemanticKind;
        JsonType = jsonType;
        Signals = signals ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Classified payload format.</summary>
    public MarkdownNativeVisualPayloadFormat Format { get; }

    /// <summary>Whether the payload appears to be JSON.</summary>
    public bool IsJson => Format == MarkdownNativeVisualPayloadFormat.JsonObject || Format == MarkdownNativeVisualPayloadFormat.JsonArray;

    /// <summary>Whether the payload appears to be a Mermaid diagram.</summary>
    public bool IsMermaid => Format == MarkdownNativeVisualPayloadFormat.Mermaid;

    /// <summary>Semantic kind declared by the fenced block.</summary>
    public string DeclaredSemanticKind { get; }

    /// <summary>Best-effort semantic kind detected from the payload, when available.</summary>
    public string? DetectedSemanticKind { get; }

    /// <summary>Best-effort value of a JSON <c>type</c> property, when available.</summary>
    public string? JsonType { get; }

    /// <summary>Small dependency-free signal map for UI hosts that need quick routing hints.</summary>
    public IReadOnlyDictionary<string, string> Signals { get; }

    /// <summary>
    /// Classifies a semantic fenced block payload without introducing a JSON parser dependency.
    /// </summary>
    public static MarkdownNativeVisualPayload Create(SemanticFencedBlock visual) {
        if (visual == null) {
            throw new ArgumentNullException(nameof(visual));
        }

        var content = visual.Content ?? string.Empty;
        var trimmed = content.Trim();
        var signals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var format = ClassifyFormat(visual, trimmed);
        if (format != MarkdownNativeVisualPayloadFormat.Unknown) {
            signals["format"] = format.ToString();
        }

        string? detectedSemanticKind = null;
        if (MarkdownJsonVisualPayloadDetector.TryDetectSemanticKind(content, out var detected)) {
            detectedSemanticKind = detected;
            signals["detectedSemanticKind"] = detected;
        }

        var jsonType = TryReadJsonStringProperty(trimmed, "type");
        if (!string.IsNullOrWhiteSpace(jsonType)) {
            signals["json.type"] = jsonType!;
        }

        AddPresenceSignal(signals, trimmed, "data");
        AddPresenceSignal(signals, trimmed, "options");
        AddPresenceSignal(signals, trimmed, "nodes");
        AddPresenceSignal(signals, trimmed, "edges");
        AddPresenceSignal(signals, trimmed, "rows");
        AddPresenceSignal(signals, trimmed, "columns");
        AddPresenceSignal(signals, trimmed, "items");

        return new MarkdownNativeVisualPayload(
            format,
            visual.SemanticKind,
            detectedSemanticKind,
            jsonType,
            signals);
    }

    private static MarkdownNativeVisualPayloadFormat ClassifyFormat(SemanticFencedBlock visual, string trimmedContent) {
        if (string.IsNullOrWhiteSpace(trimmedContent)) {
            return MarkdownNativeVisualPayloadFormat.Unknown;
        }

        if (string.Equals(visual.SemanticKind, MarkdownSemanticKinds.Mermaid, StringComparison.OrdinalIgnoreCase)
            || string.Equals(visual.Language, "mermaid", StringComparison.OrdinalIgnoreCase)) {
            return MarkdownNativeVisualPayloadFormat.Mermaid;
        }

        if (trimmedContent.Length >= 2) {
            if (trimmedContent[0] == '{' && trimmedContent[trimmedContent.Length - 1] == '}') {
                return MarkdownNativeVisualPayloadFormat.JsonObject;
            }

            if (trimmedContent[0] == '[' && trimmedContent[trimmedContent.Length - 1] == ']') {
                return MarkdownNativeVisualPayloadFormat.JsonArray;
            }
        }

        return MarkdownNativeVisualPayloadFormat.Text;
    }

    private static void AddPresenceSignal(IDictionary<string, string> signals, string payload, string propertyName) {
        if (ContainsJsonProperty(payload, propertyName)) {
            signals["json.has." + propertyName] = "true";
        }
    }

    private static bool ContainsJsonProperty(string payload, string propertyName) {
        if (string.IsNullOrWhiteSpace(payload) || string.IsNullOrWhiteSpace(propertyName)) {
            return false;
        }

        return payload.IndexOf("\"" + propertyName + "\"", StringComparison.OrdinalIgnoreCase) >= 0
               || payload.IndexOf("'" + propertyName + "'", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string? TryReadJsonStringProperty(string payload, string propertyName) {
        if (string.IsNullOrWhiteSpace(payload) || string.IsNullOrWhiteSpace(propertyName)) {
            return null;
        }

        var quotedName = "\"" + propertyName + "\"";
        var index = payload.IndexOf(quotedName, StringComparison.OrdinalIgnoreCase);
        var quoteLength = quotedName.Length;
        if (index < 0) {
            quotedName = "'" + propertyName + "'";
            index = payload.IndexOf(quotedName, StringComparison.OrdinalIgnoreCase);
            quoteLength = quotedName.Length;
        }

        if (index < 0) {
            return null;
        }

        var colon = payload.IndexOf(':', index + quoteLength);
        if (colon < 0) {
            return null;
        }

        var valueStart = colon + 1;
        while (valueStart < payload.Length && char.IsWhiteSpace(payload[valueStart])) {
            valueStart++;
        }

        if (valueStart >= payload.Length) {
            return null;
        }

        var quote = payload[valueStart];
        if (quote != '"' && quote != '\'') {
            return null;
        }

        var valueEnd = valueStart + 1;
        while (valueEnd < payload.Length) {
            if (payload[valueEnd] == quote && payload[valueEnd - 1] != '\\') {
                break;
            }

            valueEnd++;
        }

        return valueEnd > valueStart + 1 && valueEnd < payload.Length
            ? payload.Substring(valueStart + 1, valueEnd - valueStart - 1)
            : null;
    }
}
