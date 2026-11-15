using OfficeIMO.Query;

namespace OfficeIMO.Yaml;

/// <summary>A node a query selected, with the path that addresses it.</summary>
public sealed class YamlQueryMatch {
    internal YamlQueryMatch(YamlNode node, string path) {
        Node = node;
        Path = path;
    }

    /// <summary>
    /// Gets the selected node as written: a value reached through an alias is the alias's anchored node, and a mapping
    /// value written as an alias is the <see cref="YamlAlias"/>. <see cref="YamlNode.Start"/> and
    /// <see cref="YamlNode.End"/> give its source position.
    /// </summary>
    public YamlNode Node { get; }

    /// <summary>
    /// Gets the path of the node in the grammar <see cref="YamlNode.Select"/> reads, relative to where the query ran:
    /// <c>Select(Path)</c> on that node or document returns <see cref="Node"/>. The root is <c>$</c>.
    /// </summary>
    public string Path { get; }

    /// <summary>Returns the path and the node's start position.</summary>
    public override string ToString() => Node.HasSourceSpan ? Path + " at " + Node.Start : Path;
}

/// <summary>Runs <see cref="StructuredQuery"/> queries over the YAML model.</summary>
internal static class YamlQuery {
    private static readonly Adapter Instance = new Adapter();

    public static IReadOnlyList<YamlQueryMatch> Query(YamlDocument document, string query) {
        StructuredQuery parsed = StructuredQuery.Parse(query, "YAML");
        IReadOnlyList<YamlDocumentPart> parts = document.Documents;
        var start = new List<QueryMatch<YamlNode>>();
        int first = 0;
        if (parts.Count > 1 && parsed.Rooted && parsed.Steps.Count > 0 && parsed.Steps[0] is { Descendants: false, Kind: not QueryStepKind.Key } documents) {
            // A leading $[n], $[*] or $[?(...)] picks documents of a stream, as $[n] does for Select.
            first = 1;
            for (int i = 0; i < parts.Count; i++) {
                bool selected = documents.Kind switch {
                    QueryStepKind.Index => i == (documents.Index < 0 ? parts.Count + documents.Index : documents.Index),
                    QueryStepKind.Wildcard => true,
                    _ => documents.Filter!.IsTrue(Instance, parts[i].Root)
                };
                if (selected) start.Add(new QueryMatch<YamlNode>(parts[i].Root, "$[" + i.ToString(CultureInfo.InvariantCulture) + "]"));
            }
        } else if (parts.Count > 0) {
            start.Add(new QueryMatch<YamlNode>(parts[0].Root, string.Empty));
        }
        return Results(parsed.Evaluate(Instance, start, first));
    }

    public static IReadOnlyList<YamlQueryMatch> Query(YamlNode node, string query) {
        StructuredQuery parsed = StructuredQuery.Parse(query, "YAML");
        return Results(parsed.Evaluate(Instance, new List<QueryMatch<YamlNode>> { new QueryMatch<YamlNode>(node, string.Empty) }, 0));
    }

    private static IReadOnlyList<YamlQueryMatch> Results(List<QueryMatch<YamlNode>> matches) {
        if (matches.Count == 0) return Array.Empty<YamlQueryMatch>();
        var results = new YamlQueryMatch[matches.Count];
        for (int i = 0; i < results.Length; i++) results[i] = new YamlQueryMatch(matches[i].Node, matches[i].Path.Length == 0 ? "$" : matches[i].Path);
        return results;
    }

    private sealed class Adapter : IQueryAdapter<YamlNode> {
        public bool IsReference(YamlNode node) => node is YamlAlias;

        public QueryNodeKind GetKind(YamlNode node) => node.Resolved switch {
            YamlMapping => QueryNodeKind.Mapping,
            YamlSequence => QueryNodeKind.Sequence,
            _ => QueryNodeKind.Scalar
        };

        public YamlNode? GetMember(YamlNode node, string key) => YamlPath.Key(node, key);

        public int GetCount(YamlNode node) => ((YamlSequence)node.Resolved)._count;

        public YamlNode GetItem(YamlNode node, int index) => ((YamlSequence)node.Resolved).ItemAt(index);

        public void AddChildren(YamlNode node, bool written, List<QueryChild<YamlNode>> children) {
            if (written && node is YamlAlias) return;
            YamlNode resolved = node.Resolved;
            if (resolved is YamlSequence sequence) {
                for (int i = 0; i < sequence._count; i++) children.Add(new QueryChild<YamlNode>(i, sequence.ItemAt(i)));
            } else if (resolved is YamlMapping mapping) {
                if (written) {
                    for (int i = 0; i < mapping.Count; i++) {
                        if (mapping.IsMergeKeyAt(i)) continue;
                        YamlMappingEntry entry = mapping.EntryAt(i);
                        children.Add(new QueryChild<YamlNode>(KeyText(entry.Key), entry.Value));
                    }
                } else {
                    AddMembers(mapping, children, new HashSet<string>(StringComparer.Ordinal), 0);
                }
            }
        }

        /// <summary>
        /// Adds the values key steps reach: own entries other than merge keys, then merged keys no earlier entry has,
        /// in the precedence <see cref="YamlMapping.TryGetValue"/> uses.
        /// </summary>
        private static void AddMembers(YamlMapping mapping, List<QueryChild<YamlNode>> children, HashSet<string> seen, int depth) {
            if (depth > 64) return;
            List<YamlNode>? merges = null;
            for (int i = 0; i < mapping.Count; i++) {
                YamlMappingEntry entry = mapping.EntryAt(i);
                if (mapping.IsMergeKeyAt(i)) {
                    (merges ??= new List<YamlNode>()).Add(entry.Value.Resolved);
                    continue;
                }
                string key = KeyText(entry.Key);
                if (seen.Add(key)) children.Add(new QueryChild<YamlNode>(key, entry.Value));
            }
            if (merges is null) return;
            foreach (YamlNode source in merges) {
                if (source is YamlMapping merged) {
                    AddMembers(merged, children, seen, depth + 1);
                } else if (source is YamlSequence list) {
                    for (int s = 0; s < list._count; s++) {
                        if (list.ItemAt(s).Resolved is YamlMapping item) AddMembers(item, children, seen, depth + 1);
                    }
                }
            }
        }

        /// <summary>A scalar key by its decoded text; a collection key by its YAML text, as paths address it.</summary>
        private static string KeyText(YamlNode key) => key.Resolved is YamlScalar scalar ? scalar.Value : key.Resolved.ToString();

        public bool TryGetScalar(YamlNode node, out QueryScalar value) {
            if (node.Resolved is not YamlScalar scalar) {
                value = default;
                return false;
            }
            switch (scalar.Kind) {
                case YamlScalarKind.Null:
                    value = QueryScalar.Null;
                    return true;
                case YamlScalarKind.Boolean when scalar.TryGetValue(out bool boolean):
                    value = QueryScalar.FromBoolean(boolean);
                    return true;
                case YamlScalarKind.Integer or YamlScalarKind.Float when scalar.TryGetValue(out double number):
                    value = QueryScalar.FromNumber(number);
                    return true;
                default:
                    value = QueryScalar.FromString(scalar.Value);
                    return true;
            }
        }

        public string AppendKey(string path, string key) => YamlPath.AppendKey(path, key);

        public string AppendIndex(string path, int index) => YamlPath.AppendIndex(path, index);
    }
}
