using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static readonly HashSet<string> MergeRelationshipNames = new HashSet<string>(new[] {
        "name", "usemap", "for", "form", "list", "headers", "itemref", "aria-labelledby", "aria-describedby",
        "aria-controls", "aria-owns", "aria-flowto", "aria-activedescendant", "aria-details", "aria-errormessage"
    }, StringComparer.Ordinal);

    private static readonly HashSet<string> MergeResourceAttributeNames = new HashSet<string>(new[] {
        "href", "src", "poster", "cite", "longdesc", "data", "definitionURL", "srcset", "imagesrcset", "style",
        "fill", "stroke", "filter", "clip-path", "mask", "marker", "marker-start", "marker-mid", "marker-end",
        "cursor", "ping", "archive"
    }, StringComparer.Ordinal);

    private static List<(XAttribute Attribute, string Original)> CaptureMergeRelationshipAttributes(XElement root, CancellationToken token) {
        var result = new List<(XAttribute, string)>();
        foreach (XAttribute attribute in root.DescendantsAndSelf().Attributes()) {
            token.ThrowIfCancellationRequested();
            if (attribute.Name.NamespaceName.Length == 0 &&
                (MergeRelationshipNames.Contains(attribute.Name.LocalName) || MergeResourceAttributeNames.Contains(attribute.Name.LocalName)))
                result.Add((attribute, attribute.Value));
        }
        return result;
    }

    // Preserve attribute-selector truth over the second document, including foreign vocabulary and
    // ordinary name/for attributes that the content rewriter intentionally leaves unchanged.
    private sealed class MergeRelationshipSelectors {
        private static readonly char[] Separators = { ' ', '\t', '\r', '\n', '\f' };
        private readonly Dictionary<string, (MergeAttributeValues Exact, MergeAttributeValues Tokens)> _attributes =
            new Dictionary<string, (MergeAttributeValues, MergeAttributeValues)>(StringComparer.Ordinal);
        private readonly HashSet<string> _removedNames = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _changedTokenShapes = new HashSet<string>(StringComparer.Ordinal);
        private readonly CancellationToken _token;

        internal MergeRelationshipSelectors(List<(XAttribute Attribute, string Original)> attributes, CancellationToken token) {
            _token = token;
            foreach (var attribute in attributes) {
                token.ThrowIfCancellationRequested();
                string name = attribute.Attribute.Name.LocalName;
                if (!_attributes.TryGetValue(name, out var values)) {
                    values = (new MergeAttributeValues(), new MergeAttributeValues());
                    _attributes.Add(name, values);
                }
                string? current = attribute.Attribute.Document == null ? null : attribute.Attribute.Value;
                if (current == null) _removedNames.Add(name);
                values.Exact.Add(attribute.Original, current);
                if (!MergeRelationshipNames.Contains(name)) continue;
                string[] before = attribute.Original.Split(Separators, StringSplitOptions.RemoveEmptyEntries);
                string[] after = (current ?? string.Empty).Split(Separators, StringSplitOptions.RemoveEmptyEntries);
                if (before.Length != after.Length) { _changedTokenShapes.Add(name); continue; }
                for (int i = 0; i < before.Length; i++) { token.ThrowIfCancellationRequested(); values.Tokens.Add(before[i], after[i]); }
            }
        }

        internal string Rewrite(string name, string operation, string value) {
            _token.ThrowIfCancellationRequested();
            bool relationship = MergeRelationshipNames.Contains(name);
            if (!relationship && !MergeResourceAttributeNames.Contains(name))
                throw new NotSupportedException("Case-variant relationship selectors require explicit reconciliation.");
            if (operation == "~=" && (!relationship || _changedTokenShapes.Contains(name)))
                throw new NotSupportedException("Token selectors on resource values require explicit reconciliation.");
            if (operation.Length == 0) {
                if (_removedNames.Contains(name)) throw new NotSupportedException("An attribute-presence selector refers to removed document scaffolding.");
                return value;
            }
            if (!_attributes.TryGetValue(name, out var values)) return value;
            // A whitespace-containing or empty ~= operand cannot match a token before or after repair.
            if (operation == "~=" && (value.Length == 0 || value.IndexOfAny(Separators) >= 0)) return value;
            return (operation == "~=" ? values.Tokens : values.Exact).Rewrite(value);
        }
    }

    private sealed class MergeAttributeValues {
        private readonly Dictionary<string, HashSet<string>> _forward = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        private readonly Dictionary<string, HashSet<string>> _reverse = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);

        private readonly HashSet<string> _removed = new HashSet<string>(StringComparer.Ordinal);

        internal void Add(string before, string? after) {
            if (after == null) { _removed.Add(before); return; }
            Add(_forward, before, after); Add(_reverse, after, before);
        }

        internal string Rewrite(string value) {
            if (_removed.Contains(value)) throw new NotSupportedException("An attribute selector refers to removed document scaffolding.");
            string replacement = value;
            if (_forward.TryGetValue(value, out var targets)) {
                if (targets.Count != 1) throw new NotSupportedException("A relationship selector requires different replacements on different elements.");
                replacement = targets.First();
            }
            if (_reverse.TryGetValue(replacement, out var sources) && (sources.Count != 1 || !sources.Contains(value)))
                throw new NotSupportedException("A repaired relationship selector would match additional attribute values.");
            return replacement;
        }

        private static void Add(Dictionary<string, HashSet<string>> values, string key, string value) {
            if (!values.TryGetValue(key, out var entries)) { entries = new HashSet<string>(StringComparer.Ordinal); values.Add(key, entries); }
            entries.Add(value);
        }
    }
}
