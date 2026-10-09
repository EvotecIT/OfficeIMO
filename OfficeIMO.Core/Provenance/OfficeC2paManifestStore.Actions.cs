using System.Collections.Generic;

namespace OfficeIMO.Provenance;

internal static partial class OfficeC2paManifestStore {
    /// <summary>Projects the effective action fields after v2 templates and related-action inheritance.</summary>
    private static void DescribeActions(Dictionary<object, object?> assertion, string label, List<OfficeC2paAction> output, ref bool declaresGenerativeAi) {
        if (Get(assertion, "actions") is not List<object?> actions) return;
        bool versionTwo = label == "c2pa.actions.v2" || label.StartsWith("c2pa.actions.v2__", System.StringComparison.Ordinal);
        var templates = versionTwo ? Get(assertion, "templates") as List<object?> : null;
        var agents = versionTwo ? Get(assertion, "softwareAgents") as List<object?> : null;
        foreach (object? entry in actions) {
            if (entry is not Dictionary<object, object?> action) continue;
            DescribeAction(action, null, templates, agents, versionTwo, output, 0, ref declaresGenerativeAi);
        }
    }

    private static void DescribeAction(Dictionary<object, object?> action, Dictionary<object, object?>? parent,
        List<object?>? templates, List<object?>? agents, bool versionTwo, List<OfficeC2paAction> output, int depth, ref bool declaresGenerativeAi) {
        if (depth > 16) return;
        string? name = Text(Get(action, "action")) ?? (parent == null ? null : Text(Get(parent, "action")));
        if (string.IsNullOrEmpty(name)) return;
        var effective = new Dictionary<object, object?>();
        if (templates != null) {
            // Wildcards are defaults regardless of their position; matching templates override them in array order.
            foreach (string match in new[] { "*", name! }) {
                foreach (object? entry in templates) {
                    if (entry is Dictionary<object, object?> template && Text(Get(template, "action")) == match) Overlay(effective, template);
                }
            }
        }
        if (parent != null) Overlay(effective, parent);
        Overlay(effective, action);
        string? agent = Agent(Get(effective, "softwareAgent"));
        if (agent == null && agents != null && Get(effective, "softwareAgentIndex") is long index && index >= 0 && index < agents.Count) {
            agent = Agent(agents[(int)index]);
        }
        object? when = Get(effective, "when");
        if (when is OfficeCborTag tag) when = tag.Tag == 0 ? tag.Value : null;
        var record = new OfficeC2paAction(name!, agent, Text(Get(effective, "digitalSourceType")), Text(when));
        declaresGenerativeAi |= record.DigitalSourceKind is OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia
            or OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia;
        if (output.Count < MaximumDescribedActions) output.Add(record);
        if (versionTwo && Get(action, "related") is List<object?> related) {
            foreach (object? entry in related) {
                if (entry is Dictionary<object, object?> child) DescribeAction(child, effective, templates, agents, true, output, depth + 1, ref declaresGenerativeAi);
            }
        }
    }

    private static void Overlay(Dictionary<object, object?> target, Dictionary<object, object?> source) {
        // The inline and indexed representations select one agent. An explicit replacement, even invalid,
        // must not fall back to the inherited alternative and attribute the action to a different tool.
        if (source.ContainsKey("softwareAgent")) target.Remove("softwareAgentIndex");
        else if (source.ContainsKey("softwareAgentIndex")) target.Remove("softwareAgent");
        foreach (KeyValuePair<object, object?> field in source) target[field.Key] = field.Value;
    }
}
