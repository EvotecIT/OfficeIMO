using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Provenance;

internal static partial class OfficeC2paManifestStore {
    /// <summary>Projects the effective action fields after v2 templates and related-action inheritance.</summary>
    private static void DescribeActions(Dictionary<object, object?> assertion, string label, List<OfficeC2paAction> output, ref bool declaresGenerativeAi) {
        if (Get(assertion, "actions") is not List<object?> actions) return;
        bool versionTwo = label == "c2pa.actions.v2" || label.StartsWith("c2pa.actions.v2__", System.StringComparison.Ordinal);
        var templates = IndexTemplates(versionTwo ? Get(assertion, "templates") as List<object?> : null);
        var cache = new ActionProjectionCache(templates);
        var agents = versionTwo ? Get(assertion, "softwareAgents") as List<object?> : null;
        foreach (object? entry in actions) {
            if (entry is not Dictionary<object, object?> action) continue;
            DescribeAction(action, null, templates, cache, agents, versionTwo, output, 0, ref declaresGenerativeAi);
        }
    }

    private static void DescribeAction(Dictionary<object, object?> action, Dictionary<object, object?>? parent,
        Dictionary<string, Dictionary<object, object?>> templates, ActionProjectionCache cache, List<object?>? agents,
        bool versionTwo, List<OfficeC2paAction> output, int depth, ref bool declaresGenerativeAi) {
        if (depth > 16) return;
        string? name = Get(action, "action") as string ?? (parent == null ? null : Get(parent, "action") as string);
        if (string.IsNullOrEmpty(name)) return;
        var effective = new Dictionary<object, object?>();
        // Wildcards are defaults regardless of position. Each indexed template already incorporates
        // its matching entries in array order, so each action performs only two bounded overlays.
        if (cache.Wildcard != null) Overlay(effective, cache.Wildcard);
        if (name != "*" && cache.MatchingTemplate(name!, templates) is Dictionary<object, object?> matching) Overlay(effective, matching);
        if (parent != null) Overlay(effective, parent);
        Overlay(effective, action);
        string? agent = cache.Agent(Get(effective, "softwareAgent"));
        if (agent == null && agents != null && Get(effective, "softwareAgentIndex") is long index && index >= 0 && index < agents.Count) {
            agent = cache.Agent(agents[(int)index]);
        }
        object? when = Get(effective, "when");
        if (when is OfficeCborTag tag) when = tag.Tag == 0 ? tag.Value : null;
        var record = new OfficeC2paAction(cache.Text(name) ?? "", agent,
            cache.Text(Get(effective, "digitalSourceType")), cache.Text(when));
        declaresGenerativeAi |= record.DigitalSourceKind is OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia
            or OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia;
        if (output.Count < MaximumDescribedActions) output.Add(record);
        if (versionTwo && Get(action, "related") is List<object?> related) {
            foreach (object? entry in related) {
                if (entry is Dictionary<object, object?> child) DescribeAction(child, effective, templates, cache, agents, true, output, depth + 1, ref declaresGenerativeAi);
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

    private sealed class ActionProjectionCache {
        private readonly Dictionary<object, string?> _texts = new(ReferenceComparer.Instance);
        private readonly Dictionary<object, string?> _agents = new(ReferenceComparer.Instance);
        private readonly Dictionary<object, Dictionary<object, object?>?> _matches = new(ReferenceComparer.Instance);

        internal ActionProjectionCache(Dictionary<string, Dictionary<object, object?>> templates) {
            templates.TryGetValue("*", out Dictionary<object, object?>? wildcard);
            Wildcard = wildcard;
        }

        internal Dictionary<object, object?>? Wildcard { get; }

        internal Dictionary<object, object?>? MatchingTemplate(string name,
            Dictionary<string, Dictionary<object, object?>> templates) {
            if (_matches.TryGetValue(name, out Dictionary<object, object?>? match)) return match;
            templates.TryGetValue(name, out match);
            _matches.Add(name, match);
            return match;
        }

        internal string? Text(object? value) {
            if (value is not string source) return null;
            if (_texts.TryGetValue(source, out string? result)) return result;
            result = OfficeC2paManifestStore.Text(source);
            _texts.Add(source, result);
            return result;
        }

        internal string? Agent(object? value) {
            if (value is string) return Text(value);
            if (value is not Dictionary<object, object?> source) return null;
            if (_agents.TryGetValue(source, out string? result)) return result;
            string? name = Text(Get(source, "name"));
            string? version = Text(Get(source, "version"));
            result = name == null ? null : version == null ? name : name + " " + version;
            _agents.Add(source, result);
            return result;
        }
    }

    private sealed class ReferenceComparer : IEqualityComparer<object> {
        internal static readonly ReferenceComparer Instance = new();
        public new bool Equals(object? left, object? right) => ReferenceEquals(left, right);
        public int GetHashCode(object value) => RuntimeHelpers.GetHashCode(value);
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
