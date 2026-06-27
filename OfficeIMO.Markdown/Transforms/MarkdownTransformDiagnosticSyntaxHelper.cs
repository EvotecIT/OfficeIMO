namespace OfficeIMO.Markdown;

internal static class MarkdownTransformDiagnosticSyntaxHelper {
    internal static void PopulateOriginalBlockAnchor(
        MarkdownDocumentTransformDiagnostic diagnostic,
        MarkdownSyntaxNode? originalSyntaxTree) {
        PopulateBlockAnchor(
            diagnostic,
            originalSyntaxTree,
            setPath: static (item, value) => item.AffectedOriginalBlockPath = value,
            setSpan: static (item, value) => item.AffectedOriginalBlockSpan = value);
        PopulateNodeAnchor(
            diagnostic,
            originalSyntaxTree,
            setPath: static (item, value) => item.AffectedOriginalNodePath = value,
            setSpan: static (item, value) => item.AffectedOriginalNodeSpan = value);
    }

    internal static void PopulateFinalBlockAnchors(
        IReadOnlyList<MarkdownDocumentTransformDiagnostic>? diagnostics,
        MarkdownSyntaxNode? finalSyntaxTree) {
        if (diagnostics == null || finalSyntaxTree == null) {
            return;
        }

        for (var i = 0; i < diagnostics.Count; i++) {
            PopulateBlockAnchor(
                diagnostics[i],
                finalSyntaxTree,
                setPath: static (item, value) => item.AffectedFinalBlockPath = value,
                setSpan: static (item, value) => item.AffectedFinalBlockSpan = value);
            PopulateNodeAnchor(
                diagnostics[i],
                finalSyntaxTree,
                setPath: static (item, value) => item.AffectedFinalNodePath = value,
                setSpan: static (item, value) => item.AffectedFinalNodeSpan = value);
        }
    }

    internal static void PopulateOriginalChangedNodeAnchor(
        MarkdownDocumentTransformDiagnostic diagnostic,
        MarkdownSyntaxNode? beforeNode,
        MarkdownSyntaxNode? afterNode,
        MarkdownTransformSourceSpanHelper.ChangedBlockRange change) {
        if (diagnostic == null ||
            beforeNode == null ||
            afterNode == null ||
            change.CountBefore != 1 ||
            change.CountAfter != 1) {
            return;
        }

        var changedPath = FindChangedNodePath(beforeNode, afterNode);
        if (changedPath.Count == 0) {
            return;
        }

        var changedNode = changedPath[changedPath.Count - 1];
        if (!changedNode.SourceSpan.HasValue) {
            return;
        }

        diagnostic.AffectedOriginalNodePath = "Document > " + string.Join(" > ", changedPath.Select(FormatPathSegment));
        diagnostic.AffectedOriginalNodeSpan = changedNode.SourceSpan;
    }

    private static void PopulateBlockAnchor(
        MarkdownDocumentTransformDiagnostic diagnostic,
        MarkdownSyntaxNode? syntaxTree,
        Action<MarkdownDocumentTransformDiagnostic, string?> setPath,
        Action<MarkdownDocumentTransformDiagnostic, MarkdownSourceSpan?> setSpan) {
        if (diagnostic == null || syntaxTree == null || !diagnostic.AffectedSourceSpan.HasValue) {
            return;
        }

        var blockPath = FindBlockPath(syntaxTree, diagnostic.AffectedSourceSpan.Value);
        if (blockPath.Count == 0) {
            return;
        }

        setPath(diagnostic, string.Join(" > ", blockPath.Select(FormatPathSegment)));
        setSpan(diagnostic, blockPath[blockPath.Count - 1].SourceSpan);
    }

    private static void PopulateNodeAnchor(
        MarkdownDocumentTransformDiagnostic diagnostic,
        MarkdownSyntaxNode? syntaxTree,
        Action<MarkdownDocumentTransformDiagnostic, string?> setPath,
        Action<MarkdownDocumentTransformDiagnostic, MarkdownSourceSpan?> setSpan) {
        if (diagnostic == null || syntaxTree == null || !diagnostic.AffectedSourceSpan.HasValue) {
            return;
        }

        var nodePath = FindNodePath(syntaxTree, diagnostic.AffectedSourceSpan.Value);
        if (nodePath.Count == 0) {
            return;
        }

        setPath(diagnostic, string.Join(" > ", nodePath.Select(FormatPathSegment)));
        setSpan(diagnostic, nodePath[nodePath.Count - 1].SourceSpan);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> FindBlockPath(MarkdownSyntaxNode syntaxTree, MarkdownSourceSpan span) {
        var path = syntaxTree.FindNodePathContainingSpan(span);
        if (path.Count == 0) {
            path = syntaxTree.FindNodePathOverlappingSpan(span);
        }

        if (path.Count == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var blockPath = new List<MarkdownSyntaxNode>(path.Count);
        for (var i = 0; i < path.Count; i++) {
            if (path[i].IsBlockLike) {
                blockPath.Add(path[i]);
            }
        }

        return blockPath;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> FindNodePath(MarkdownSyntaxNode syntaxTree, MarkdownSourceSpan span) {
        var path = syntaxTree.FindNodePathContainingSpan(span);
        if (path.Count == 0) {
            path = syntaxTree.FindNodePathOverlappingSpan(span);
        }

        return path;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> FindChangedNodePath(
        MarkdownSyntaxNode beforeNode,
        MarkdownSyntaxNode afterNode) {
        if (beforeNode == null || afterNode == null) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        if (NodeFingerprint(beforeNode) == NodeFingerprint(afterNode)) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        if (beforeNode.Kind != afterNode.Kind ||
            !string.Equals(beforeNode.CustomKind, afterNode.CustomKind, StringComparison.Ordinal)) {
            return new[] { beforeNode };
        }

        var childChange = ComputeChildChange(beforeNode.Children, afterNode.Children);
        if (childChange.CountBefore == 1 && childChange.CountAfter == 1) {
            var beforeChild = beforeNode.Children[childChange.StartBefore];
            var afterChild = afterNode.Children[childChange.StartAfter];
            var childPath = FindChangedNodePath(beforeChild, afterChild);
            if (childPath.Count > 0) {
                var path = new List<MarkdownSyntaxNode>(childPath.Count + 1) {
                    beforeNode
                };
                path.AddRange(childPath);
                return path;
            }

            return new[] { beforeNode, beforeChild };
        }

        return new[] { beforeNode };
    }

    private static MarkdownTransformSourceSpanHelper.ChangedBlockRange ComputeChildChange(
        IReadOnlyList<MarkdownSyntaxNode> before,
        IReadOnlyList<MarkdownSyntaxNode> after) {
        var beforeFingerprints = new string[before.Count];
        for (var i = 0; i < before.Count; i++) {
            beforeFingerprints[i] = NodeFingerprint(before[i]);
        }

        var afterFingerprints = new string[after.Count];
        for (var i = 0; i < after.Count; i++) {
            afterFingerprints[i] = NodeFingerprint(after[i]);
        }

        return MarkdownTransformSourceSpanHelper.ComputeChangedRange(beforeFingerprints, afterFingerprints);
    }

    private static string NodeFingerprint(MarkdownSyntaxNode node) {
        if (node == null) {
            return string.Empty;
        }

        return string.Concat(
            node.Kind.ToString(),
            "\n",
            node.CustomKind ?? string.Empty,
            "\n",
            node.Literal ?? string.Empty,
            "\n",
            string.Join("\n", node.Children.Select(NodeFingerprint)));
    }

    private static string FormatPathSegment(MarkdownSyntaxNode node) {
        if (string.IsNullOrWhiteSpace(node.CustomKind)) {
            return node.Kind.ToString();
        }

        return node.Kind + "(" + node.CustomKind + ")";
    }
}
