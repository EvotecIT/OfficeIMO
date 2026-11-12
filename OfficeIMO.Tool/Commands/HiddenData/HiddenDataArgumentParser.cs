namespace OfficeIMO.Tool.Commands.HiddenData;

/// <summary>Parses the hidden-data option values shared by the CLI and the MCP tools.</summary>
internal static class HiddenDataArgumentParser {
    /// <summary>Parses kind names such as <c>default</c>, <c>all</c>, <c>provenance</c> or <c>hidden-content</c>.</summary>
    internal static OfficeHiddenDataKind ParseKinds(IEnumerable<string> values) {
        OfficeHiddenDataKind kinds = OfficeHiddenDataKind.None;
        foreach (string value in values) {
            string name = value.Trim();
            if (name.Length == 0) continue;
            string normalized = name.Replace("-", string.Empty, StringComparison.Ordinal);
            if (normalized.Any(char.IsDigit) ||
                !Enum.TryParse(normalized, ignoreCase: true, out OfficeHiddenDataKind kind) || kind == OfficeHiddenDataKind.None) {
                throw new ArgumentException("Unknown hidden-data kind '" + name + "'.");
            }
            kinds |= kind;
        }
        if (kinds == OfficeHiddenDataKind.None) throw new ArgumentException("At least one hidden-data kind is required.");
        return kinds;
    }

    internal static OfficeTrackedChangesAction ParseTrackedChanges(string value) => value.Trim().ToLowerInvariant() switch {
        "keep" => OfficeTrackedChangesAction.Keep,
        "accept" => OfficeTrackedChangesAction.AcceptAll,
        "reject" => OfficeTrackedChangesAction.RejectAll,
        _ => throw new ArgumentException("Unknown tracked-changes action '" + value + "'; use keep, accept or reject.")
    };

    /// <summary>Keeps the named provenance carriers when provenance is removed: <c>c2pa</c>, <c>external</c> or <c>ai-source</c>.</summary>
    internal static void ApplyKeptProvenance(OfficeHiddenDataOptions options, IEnumerable<string> values) {
        foreach (string value in values) {
            switch (value.Trim().ToLowerInvariant()) {
                case "":
                    break;
                case "c2pa":
                    options.Provenance.RemoveC2paManifests = false;
                    break;
                case "external":
                    options.Provenance.RemoveExternalC2paReferences = false;
                    break;
                case "ai-source":
                    options.Provenance.RemoveAiSourceMetadata = false;
                    break;
                default:
                    throw new ArgumentException("Unknown provenance carrier '" + value + "'; use c2pa, external or ai-source.");
            }
        }
    }
}
