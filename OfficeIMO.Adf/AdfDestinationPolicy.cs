using System.Threading;

namespace OfficeIMO.Adf;

/// <summary>Caller-defined node and mark support for a named destination product and version.</summary>
/// <remarks>
/// This policy restricts the selected structural or full-schema contract; it does not replace it.
/// A null list imposes no restriction for that category. An empty list permits none.
/// The node list applies to content nodes, including nested content, rather than the document root.
/// Lists are copied at construction. Validation does not alter the document or contact the destination.
/// </remarks>
public sealed class AdfDestinationPolicy {
    private readonly HashSet<string>? _nodes;
    private readonly HashSet<string>? _marks;

    /// <summary>Creates a destination policy from the supported content node and mark type names.</summary>
    /// <param name="name">Caller-owned product/version label, such as Jira Cloud with a qualification date.</param>
    /// <param name="allowedNodeTypes">Supported content node types, compared case-sensitively, or null to allow any type.</param>
    /// <param name="allowedMarkTypes">Supported mark types, compared case-sensitively, or null to allow any type.</param>
    public AdfDestinationPolicy(string name, IEnumerable<string>? allowedNodeTypes = null, IEnumerable<string>? allowedMarkTypes = null) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A destination product/version name is required.", nameof(name));
        Name = name;
        AllowedNodeTypes = CopyTypes(allowedNodeTypes, nameof(allowedNodeTypes), out _nodes);
        AllowedMarkTypes = CopyTypes(allowedMarkTypes, nameof(allowedMarkTypes), out _marks);
    }

    /// <summary>Gets the caller-owned destination product/version label used in diagnostics.</summary>
    public string Name { get; }
    /// <summary>Gets the immutable supported content node names, or null when nodes are unrestricted.</summary>
    public IReadOnlyList<string>? AllowedNodeTypes { get; }
    /// <summary>Gets the immutable supported mark names, or null when marks are unrestricted.</summary>
    public IReadOnlyList<string>? AllowedMarkTypes { get; }

    private static IReadOnlyList<string>? CopyTypes(IEnumerable<string>? types, string parameter, out HashSet<string>? lookup) {
        lookup = null;
        if (types == null) return null;
        var names = new List<string>();
        var distinct = new HashSet<string>(StringComparer.Ordinal);
        foreach (string type in types) {
            if (string.IsNullOrWhiteSpace(type)) throw new ArgumentException("Destination type names must be nonempty.", parameter);
            if (distinct.Add(type)) names.Add(type);
        }
        lookup = distinct;
        return Array.AsReadOnly(names.ToArray());
    }

    internal AdfValidationResult Validate(AdfDocument document, AdfValidationResult baseline, CancellationToken cancellationToken) {
        // The baseline reports unsafe graphs. Do not start another traversal through one.
        if (baseline.Issues.Any(issue => issue.Code == "ADF_CYCLIC_CONTENT" || issue.Code == "ADF_NULL_NODE" ||
            issue.Code == "ADF_NULL_MARK" || issue.Code == "ADF_CONTENT_DEPTH_EXCEEDED" || issue.Code == "ADF_NODE_LIMIT_EXCEEDED")) return baseline;
        var issues = baseline.Issues.ToList();
        var pending = new Stack<Frame>();
        pending.Push(new Frame(document.ContentItems, "$"));
        int destinationIssues = 0;
        while (pending.Count > 0 && destinationIssues < 1000) {
            cancellationToken.ThrowIfCancellationRequested();
            Frame frame = pending.Peek();
            if (frame.Index >= frame.Children.Count) { pending.Pop(); continue; }
            int index = frame.Index++;
            AdfNode node = frame.Children[index];
            string path = frame.Path + ".content[" + index + "]";
            if (_nodes != null && !_nodes.Contains(node.Type)) {
                issues.Add(Error("ADF_DESTINATION_NODE", path + ".type", "node", node.Type));
                destinationIssues++;
            }
            for (int mark = 0; mark < node.MarkItems.Count && destinationIssues < 1000; mark++) {
                cancellationToken.ThrowIfCancellationRequested();
                string type = node.MarkItems[mark].Type;
                if (_marks != null && !_marks.Contains(type)) {
                    issues.Add(Error("ADF_DESTINATION_MARK", path + ".marks[" + mark + "].type", "mark", type));
                    destinationIssues++;
                }
            }
            if (node.ContentItems.Count > 0) pending.Push(new Frame(node.ContentItems, path));
        }
        return new AdfValidationResult(issues);
    }

    private AdfValidationIssue Error(string code, string path, string category, string type) =>
        new AdfValidationIssue(code, path, "Destination '" + Name + "' does not permit " + category + " type '" + type + "'.", AdfValidationSeverity.Error);

    private sealed class Frame {
        internal Frame(IReadOnlyList<AdfNode> children, string path) { Children = children; Path = path; }
        internal IReadOnlyList<AdfNode> Children { get; }
        internal string Path { get; }
        internal int Index { get; set; }
    }
}
