using System.Threading;
using OfficeIMO.Html;

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

    // Preserve attribute-selector truth over each source document, including foreign vocabulary and
    // ordinary name/for attributes that the content rewriter intentionally leaves unchanged.
    private sealed class MergeRelationshipSelectors {
        internal static readonly char[] Separators = { ' ', '\t', '\r', '\n', '\f' };
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

        internal HtmlCssAttributeSelectorEdit Rewrite(string name, string operation, string value) {
            _token.ThrowIfCancellationRequested();
            bool relationship = MergeRelationshipNames.Contains(name);
            if (!relationship && !MergeResourceAttributeNames.Contains(name))
                throw new NotSupportedException("Case-variant relationship selectors require explicit reconciliation.");
            if (operation.Length == 0) {
                if (_removedNames.Contains(name)) throw new NotSupportedException("An attribute-presence selector refers to removed document scaffolding.");
                return HtmlCssAttributeSelectorEdit.Operand(value);
            }
            if (!_attributes.TryGetValue(name, out var values)) return HtmlCssAttributeSelectorEdit.Operand(value);
            // A whitespace-containing or empty ~= operand cannot match a token before or after repair.
            if (operation == "~=" && (value.Length == 0 || value.IndexOfAny(Separators) >= 0)) return HtmlCssAttributeSelectorEdit.Operand(value);
            if (operation == "=" && values.Exact.TryRewrite(value, out string replacement) ||
                operation == "~=" && relationship && !_changedTokenShapes.Contains(name) && values.Tokens.TryRewrite(value, out replacement))
                return HtmlCssAttributeSelectorEdit.Operand(replacement);
            return values.Exact.Expand(operation, value, _token);
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

        internal bool TryRewrite(string value, out string replacement) {
            replacement = value;
            if (_removed.Contains(value)) return false;
            if (_forward.TryGetValue(value, out var targets)) {
                if (targets.Count != 1) return false;
                replacement = targets.First();
            }
            if (_reverse.TryGetValue(replacement, out var sources) && (sources.Count != 1 || !sources.Contains(value)))
                return false;
            return true;
        }

        internal HtmlCssAttributeSelectorEdit Expand(string operation, string value, CancellationToken token) {
            bool Matches(string candidate) => operation switch {
                "=" => candidate == value,
                "~=" => value.Length != 0 && value.IndexOfAny(MergeRelationshipSelectors.Separators) < 0 &&
                    candidate.Split(MergeRelationshipSelectors.Separators, StringSplitOptions.RemoveEmptyEntries).Contains(value, StringComparer.Ordinal),
                "|=" => candidate == value || candidate.StartsWith(value + "-", StringComparison.Ordinal),
                "^=" => value.Length != 0 && candidate.StartsWith(value, StringComparison.Ordinal),
                "$=" => value.Length != 0 && candidate.EndsWith(value, StringComparison.Ordinal),
                "*=" => value.Length != 0 && candidate.IndexOf(value, StringComparison.Ordinal) >= 0,
                _ => throw new NotSupportedException("Unsupported attribute comparison.")
            };
            foreach (string removed in _removed) {
                token.ThrowIfCancellationRequested();
                if (Matches(removed)) throw new NotSupportedException("An attribute selector refers to removed document scaffolding.");
            }
            bool changed = false;
            var selected = new HashSet<string>(StringComparer.Ordinal);
            foreach (var entry in _forward) {
                token.ThrowIfCancellationRequested();
                bool before = Matches(entry.Key);
                foreach (string target in entry.Value) {
                    token.ThrowIfCancellationRequested();
                    changed |= before != Matches(target);
                    if (before) selected.Add(target);
                }
            }
            if (!changed) return HtmlCssAttributeSelectorEdit.Operand(value);
            if (selected.Count > HtmlCssAttributeSelectorEdit.MaximumAlternatives)
                throw new NotSupportedException("Attribute selector expansion exceeds 256 exact alternatives.");
            foreach (string target in selected) {
                token.ThrowIfCancellationRequested();
                foreach (string source in _reverse[target]) {
                    token.ThrowIfCancellationRequested();
                    if (!Matches(source)) throw new NotSupportedException("A repaired selector cannot distinguish matching and nonmatching source attributes.");
                }
            }
            return HtmlCssAttributeSelectorEdit.Exact(selected.OrderBy(item => item, StringComparer.Ordinal).ToArray());
        }

        private static void Add(Dictionary<string, HashSet<string>> values, string key, string value) {
            if (!values.TryGetValue(key, out var entries)) { entries = new HashSet<string>(StringComparer.Ordinal); values.Add(key, entries); }
            entries.Add(value);
        }
    }
}
