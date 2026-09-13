using System.Globalization;
using System.Text;

namespace OfficeIMO.Workflows;

public static partial class OfficeProvenanceWorkflowCatalog {
    /// <summary>Serializes the OfficeIMO-owned provenance capability contract as deterministic JSON.</summary>
    public static string ToJson() {
        var output = new StringBuilder();
        output.Append("{\n  \"id\":\"").Append(EscapeJson(Id)).Append("\",\n  \"schemaVersion\":")
            .Append(SchemaVersion.ToString(CultureInfo.InvariantCulture)).Append(",\n  \"capabilities\":[\n");
        for (int capabilityIndex = 0; capabilityIndex < All.Count; capabilityIndex++) {
            OfficeProvenanceWorkflowCapability capability = All[capabilityIndex];
            output.Append("    {\n")
                .Append("      \"id\":\"").Append(EscapeJson(capability.Id)).Append("\",\n")
                .Append("      \"label\":\"").Append(EscapeJson(capability.Label)).Append("\",\n")
                .Append("      \"ownerPackage\":\"").Append(EscapeJson(capability.OwnerPackage)).Append("\",\n")
                .Append("      \"canInspect\":").Append(capability.CanInspect ? "true" : "false").Append(",\n")
                .Append("      \"canAssess\":").Append(capability.CanAssess ? "true" : "false").Append(",\n")
                .Append("      \"canRemove\":").Append(capability.CanRemove ? "true" : "false").Append(",\n")
                .Append("      \"browserLabel\":");
            if (capability.BrowserLabel is null) output.Append("null");
            else output.Append('"').Append(EscapeJson(capability.BrowserLabel)).Append('"');
            output.Append(",\n      \"notes\":\"").Append(EscapeJson(capability.Notes)).Append("\",\n")
                .Append("      \"formats\":[\n");
            for (int formatIndex = 0; formatIndex < capability.Formats.Count; formatIndex++) {
                OfficeProvenanceWorkflowFormat format = capability.Formats[formatIndex];
                output.Append("        {\"extension\":\"").Append(EscapeJson(format.Extension))
                    .Append("\",\"assetFormats\":[")
                    .Append(string.Join(',', format.AssetFormats.Select(static item => "\"" + item + "\"")))
                    .Append("],\"memoryOnlyAvailable\":").Append(format.MemoryOnlyAvailable ? "true" : "false")
                    .Append(",\"browserAvailable\":").Append(format.BrowserAvailable ? "true" : "false")
                    .Append('}');
                if (formatIndex + 1 < capability.Formats.Count) output.Append(',');
                output.Append('\n');
            }
            output.Append("      ]\n    }");
            if (capabilityIndex + 1 < All.Count) output.Append(',');
            output.Append('\n');
        }
        return output.Append("  ]\n}").ToString();
    }

    /// <summary>Formats the OfficeIMO-owned provenance capability contract as deterministic Markdown.</summary>
    public static string ToMarkdown() {
        var output = new StringBuilder();
        output.Append("# ").Append(Id).Append(" capability contract\n\nSchema version: ")
            .Append(SchemaVersion.ToString(CultureInfo.InvariantCulture))
            .Append("\n\nOnly formats with a named OfficeIMO owner appear in this contract. Memory-only and browser support require separate qualification.\n\n")
            .Append("| Capability | Extension | Structural format | Owner | Inspect | Assess | Remove | Memory-only | Browser | Boundary |\n")
            .Append("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |\n");
        foreach (OfficeProvenanceWorkflowCapability capability in All) {
            foreach (OfficeProvenanceWorkflowFormat format in capability.Formats) {
                output.Append("| ").Append(EscapeMarkdown(capability.Label))
                    .Append(" | `").Append(EscapeMarkdown(format.Extension)).Append('`')
                    .Append(" | ").Append(string.Join(", ", format.AssetFormats))
                    .Append(" | `").Append(EscapeMarkdown(capability.OwnerPackage)).Append('`')
                    .Append(" | ").Append(capability.CanInspect ? "Yes" : "No")
                    .Append(" | ").Append(capability.CanAssess ? "Yes" : "No")
                    .Append(" | ").Append(capability.CanRemove ? "Yes" : "No")
                    .Append(" | ").Append(format.MemoryOnlyAvailable ? "Yes" : "No")
                    .Append(" | ").Append(format.BrowserAvailable ? "Yes" : "No")
                    .Append(" | ").Append(EscapeMarkdown(capability.Notes)).Append(" |\n");
            }
        }
        return output.ToString();
    }

    private static string EscapeJson(string value) {
        var escaped = new StringBuilder(value.Length + 8);
        foreach (char character in value) {
            switch (character) {
                case '"': escaped.Append("\\\""); break;
                case '\\': escaped.Append("\\\\"); break;
                case '\b': escaped.Append("\\b"); break;
                case '\f': escaped.Append("\\f"); break;
                case '\n': escaped.Append("\\n"); break;
                case '\r': escaped.Append("\\r"); break;
                case '\t': escaped.Append("\\t"); break;
                default:
                    if (character < ' ') escaped.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                    else escaped.Append(character);
                    break;
            }
        }
        return escaped.ToString();
    }

    private static string EscapeMarkdown(string value) => value
        .Replace("\\", "\\\\")
        .Replace("|", "\\|")
        .Replace("\r", " ")
        .Replace("\n", " ");
}
