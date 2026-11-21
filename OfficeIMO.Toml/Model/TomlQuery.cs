using OfficeIMO.Query;

namespace OfficeIMO.Toml;

/// <summary>A node a query selected, with the path that addresses it.</summary>
public sealed class TomlQueryMatch {
    internal TomlQueryMatch(TomlNode node, string path) {
        Node = node;
        Path = path;
    }

    /// <summary>Gets the selected node; <see cref="TomlNode.Start"/> and <see cref="TomlNode.End"/> give its source position.</summary>
    public TomlNode Node { get; }

    /// <summary>
    /// Gets the path of the node in the grammar <see cref="TomlNode.Select"/> reads, relative to where the query ran:
    /// <c>Select(Path)</c> on that node or document returns <see cref="Node"/>. The root is <c>$</c>.
    /// </summary>
    public string Path { get; }

    /// <summary>Returns the path and the node's start position.</summary>
    public override string ToString() => Node.IsSource ? Path + " at " + Node.Start : Path;
}

/// <summary>Runs <see cref="StructuredQuery"/> queries over the TOML model.</summary>
internal static class TomlQuery {
    private static readonly Adapter Instance = new Adapter();

    public static IReadOnlyList<TomlQueryMatch> Query(TomlNode node, string query) {
        StructuredQuery parsed = StructuredQuery.Parse(query, "TOML");
        List<QueryMatch<TomlNode>> matches = parsed.Evaluate(Instance, new List<QueryMatch<TomlNode>> { new QueryMatch<TomlNode>(node, string.Empty) }, 0);
        if (matches.Count == 0) return Array.Empty<TomlQueryMatch>();
        var results = new TomlQueryMatch[matches.Count];
        for (int i = 0; i < results.Length; i++) results[i] = new TomlQueryMatch(matches[i].Node, matches[i].Path.Length == 0 ? "$" : matches[i].Path);
        return results;
    }

    private sealed class Adapter : IQueryAdapter<TomlNode> {
        public bool IsReference(TomlNode node) => false;

        public QueryNodeKind GetKind(TomlNode node) => node switch {
            TomlTable => QueryNodeKind.Mapping,
            TomlArray or TomlTableArray => QueryNodeKind.Sequence,
            _ => QueryNodeKind.Scalar
        };

        public TomlNode? GetMember(TomlNode node, string key) => node is TomlTable table && table.TryGetValue(key, out TomlNode? value) ? value : null;

        public int GetCount(TomlNode node) => node is TomlArray array ? array.Count : ((TomlTableArray)node).Count;

        public TomlNode GetItem(TomlNode node, int index) => node is TomlArray array ? array[index] : ((TomlTableArray)node)[index];

        public void AddChildren(TomlNode node, bool written, List<QueryChild<TomlNode>> children) {
            switch (node) {
                case TomlTable table:
                    foreach (KeyValuePair<string, TomlNode> entry in table) children.Add(new QueryChild<TomlNode>(entry.Key, entry.Value));
                    break;
                case TomlArray array:
                    for (int i = 0; i < array.Count; i++) children.Add(new QueryChild<TomlNode>(i, array[i]));
                    break;
                case TomlTableArray tables:
                    for (int i = 0; i < tables.Count; i++) children.Add(new QueryChild<TomlNode>(i, tables[i]));
                    break;
            }
        }

        public bool TryGetScalar(TomlNode node, out QueryScalar value) {
            if (node is not TomlValue scalar) {
                value = default;
                return false;
            }
            value = scalar.Value switch {
                string text => QueryScalar.FromString(text),
                long integer => QueryScalar.FromNumber(integer),
                double number => QueryScalar.FromNumber(number),
                bool boolean => QueryScalar.FromBoolean(boolean),
                object other => QueryScalar.FromString(Convert.ToString(other, CultureInfo.InvariantCulture) ?? string.Empty)
            };
            return true;
        }

        public string AppendKey(string path, string key) => TomlPath.AppendKey(path, key);

        public string AppendIndex(string path, int index) => TomlPath.AppendIndex(path, index);
    }
}
