namespace OfficeIMO.Markdown;

/// <summary>
/// Lightweight detector for visual JSON payloads that should be treated as semantic fenced blocks.
/// </summary>
public static class MarkdownJsonVisualPayloadDetector {
    /// <summary>
    /// Best-effort detection of a known visual semantic kind from a JSON-like payload.
    /// </summary>
    /// <param name="candidate">Candidate payload text.</param>
    /// <param name="semanticKind">Detected semantic kind when successful.</param>
    /// <returns><see langword="true"/> when a known semantic kind was detected.</returns>
    public static bool TryDetectSemanticKind(string? candidate, out string semanticKind) {
        semanticKind = string.Empty;
        if (string.IsNullOrWhiteSpace(candidate)) {
            return false;
        }

        var normalized = (candidate ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        if (!LooksLikeJsonObject(normalized)) {
            return false;
        }

        if (LooksLikeNetworkPayload(normalized)) {
            semanticKind = MarkdownSemanticKinds.Network;
            return true;
        }

        if (LooksLikeDataViewPayload(normalized)) {
            semanticKind = MarkdownSemanticKinds.DataView;
            return true;
        }

        if (LooksLikeChartPayload(normalized)) {
            semanticKind = MarkdownSemanticKinds.Chart;
            return true;
        }

        return false;
    }

    private static bool LooksLikeJsonObject(string candidate) {
        var trimmed = candidate.Trim();
        return trimmed.Length >= 2
               && trimmed[0] == '{'
               && trimmed[trimmed.Length - 1] == '}';
    }

    private static bool LooksLikeNetworkPayload(string candidate) {
        return ContainsJsonProperty(candidate, "nodes")
               && ContainsJsonProperty(candidate, "edges");
    }

    private static bool LooksLikeDataViewPayload(string candidate) {
        return ContainsToken(candidate, "dataview")
               || ContainsJsonProperty(candidate, "rows")
               || ContainsJsonProperty(candidate, "records")
               || ContainsJsonProperty(candidate, "items")
               || ContainsJsonProperty(candidate, "headers")
               || ContainsJsonProperty(candidate, "columns");
    }

    private static bool LooksLikeChartPayload(string candidate) {
        return ContainsJsonProperty(candidate, "type")
               && (ContainsJsonProperty(candidate, "data") || ContainsJsonProperty(candidate, "options"));
    }

    private static bool ContainsJsonProperty(string candidate, string propertyName) {
        if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(propertyName)) {
            return false;
        }

        return candidate.IndexOf("\"" + propertyName + "\"", StringComparison.OrdinalIgnoreCase) >= 0
               || candidate.IndexOf("'" + propertyName + "'", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool ContainsToken(string candidate, string token) {
        return !string.IsNullOrWhiteSpace(candidate)
               && !string.IsNullOrWhiteSpace(token)
               && candidate.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
    }
}
