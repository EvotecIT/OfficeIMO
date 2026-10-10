using System.Collections.Generic;

namespace OfficeIMO.Provenance;

internal static partial class OfficeC2paManifestStore {
    /// <summary>Projects the effective action fields after v2 templates and related-action inheritance.</summary>
    private static void DescribeActions(Dictionary<object, object?> assertion, string label, List<OfficeC2paAction> output, ref bool declaresGenerativeAi) {
        if (Get(assertion, "actions") is not List<object?> actions) return;
        bool versionTwo = label == "c2pa.actions.v2" || label.StartsWith("c2pa.actions.v2__", System.StringComparison.Ordinal);
        var templates = IndexTemplates(versionTwo ? Get(assertion, "templates") as List<object?> : null);
        var agents = versionTwo ? Get(assertion, "softwareAgents") as List<object?> : null;
        foreach (object? entry in actions) {
            if (entry is not Dictionary<object, object?> action) continue;
            DescribeAction(action, null, templates, agents, versionTwo, output, 0, ref declaresGenerativeAi);
        }
    }

    private static void DescribeAction(Dictionary<object, object?> action, Dictionary<object, object?>? parent,
        Dictionary<string, Dictionary<object, object?>> templates, List<object?>? agents, bool versionTwo, List<OfficeC2paAction> output, int depth, ref bool declaresGenerativeAi) {
        if (depth > 16) return;
        string? name = Get(action, "action") as string ?? (parent == null ? null : Get(parent, "action") as string);
        if (string.IsNullOrEmpty(name)) return;
        var effective = new Dictionary<object, object?>();
        // Wildcards are defaults regardless of position. Each indexed template already incorporates
        // its matching entries in array order, so each action performs only two bounded overlays.
        if (templates.TryGetValue("*", out var wildcard)) Overlay(effective, wildcard);
        if (name != "*" && templates.TryGetValue(name!, out var matching)) Overlay(effective, matching);
        if (parent != null) Overlay(effective, parent);
        Overlay(effective, action);
        string? agent = Agent(Get(effective, "softwareAgent"));
        if (agent == null && agents != null && Get(effective, "softwareAgentIndex") is long index && index >= 0 && index < agents.Count) {
            agent = Agent(agents[(int)index]);
        }
        object? when = Get(effective, "when");
        if (when is OfficeCborTag tag) when = tag.Tag == 0 ? tag.Value : null;
        var record = new OfficeC2paAction(Text(name) ?? "", agent, Text(Get(effective, "digitalSourceType")), Text(when));
        declaresGenerativeAi |= record.DigitalSourceKind is OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia
            or OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia;
        if (output.Count < MaximumDescribedActions) output.Add(record);
        if (versionTwo && Get(action, "related") is List<object?> related) {
            foreach (object? entry in related) {
                if (entry is Dictionary<object, object?> child) DescribeAction(child, effective, templates, agents, true, output, depth + 1, ref declaresGenerativeAi);
            }
        }
    }

    private static Dictionary<string, Dictionary<object, object?>> IndexTemplates(List<object?>? templates) {
        var index = new Dictionary<string, Dictionary<object, object?>>(System.StringComparer.Ordinal);
        if (templates == null) return index;
        foreach (object? entry in templates) {
            if (entry is not Dictionary<object, object?> template || Get(template, "action") is not string name) continue;
            if (!index.TryGetValue(name, out var effective)) index[name] = effective = new Dictionary<object, object?>();
            Overlay(effective, template);
        }
        return index;
    }

    private static readonly string[] ActionFields = { "action", "softwareAgent", "softwareAgentIndex", "digitalSourceType", "when" };

    private static void Overlay(Dictionary<object, object?> target, Dictionary<object, object?> source) {
        // The inline and indexed representations select one agent. An explicit replacement, even invalid,
        // must not fall back to the inherited alternative and attribute the action to a different tool.
        if (source.ContainsKey("softwareAgent")) target.Remove("softwareAgentIndex");
        else if (source.ContainsKey("softwareAgentIndex")) target.Remove("softwareAgent");
        // Only these fields contribute to the summary. Copying unrelated maps or wide wildcard
        // defaults for every action would reintroduce work amplification despite template indexing.
        foreach (string field in ActionFields) if (source.TryGetValue(field, out object? value)) target[field] = value;
    }
}
