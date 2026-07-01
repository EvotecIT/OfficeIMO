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

    internal static void PopulateChangedNodeAnchors(
        IReadOnlyList<MarkdownDocumentTransformDiagnostic>? diagnostics,
        MarkdownSyntaxNode? originalSyntaxTree,
        MarkdownSyntaxNode? finalSyntaxTree) {
        if (diagnostics == null || originalSyntaxTree == null || finalSyntaxTree == null) {
            return;
        }

        for (var i = 0; i < diagnostics.Count; i++) {
            PopulateChangedNodeAnchors(diagnostics[i], originalSyntaxTree, finalSyntaxTree);
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

        var changedPaths = FindChangedNodePaths(beforeNode, afterNode);
        if (changedPaths == null) {
            return;
        }

        var changedPath = changedPaths.OriginalPath;
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

    private static void PopulateChangedNodeAnchors(
        MarkdownDocumentTransformDiagnostic diagnostic,
        MarkdownSyntaxNode originalSyntaxTree,
        MarkdownSyntaxNode finalSyntaxTree) {
        if (diagnostic == null ||
            originalSyntaxTree == null ||
            finalSyntaxTree == null ||
            !diagnostic.AffectedSourceSpan.HasValue ||
            diagnostic.ChangedBlockCountBefore != 1 ||
            diagnostic.ChangedBlockCountAfter != 1) {
            return;
        }

        var originalBlockPath = FindBlockPath(originalSyntaxTree, diagnostic.AffectedSourceSpan.Value);
        var finalBlockPath = FindBlockPath(finalSyntaxTree, diagnostic.AffectedSourceSpan.Value);
        if (originalBlockPath.Count == 0 || finalBlockPath.Count == 0) {
            return;
        }

        var changedPaths = FindChangedNodePaths(
            originalBlockPath[originalBlockPath.Count - 1],
            finalBlockPath[finalBlockPath.Count - 1]);
        if (changedPaths == null) {
            return;
        }

        ApplyChangedNodePath(
            diagnostic,
            originalBlockPath,
            changedPaths.OriginalPath,
            isOriginal: true);
        ApplyChangedNodePath(
            diagnostic,
            finalBlockPath,
            changedPaths.FinalPath,
            isOriginal: false);
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

    private static void ApplyChangedNodePath(
        MarkdownDocumentTransformDiagnostic diagnostic,
        IReadOnlyList<MarkdownSyntaxNode> blockPath,
        IReadOnlyList<MarkdownSyntaxNode> changedPath,
        bool isOriginal) {
        if (diagnostic == null ||
            blockPath == null ||
            changedPath == null ||
            blockPath.Count == 0 ||
            changedPath.Count == 0) {
            return;
        }

        var changedNode = changedPath[changedPath.Count - 1];
        if (!changedNode.SourceSpan.HasValue) {
            return;
        }

        var fullPath = new List<MarkdownSyntaxNode>(blockPath.Count + changedPath.Count);
        for (var i = 0; i < blockPath.Count - 1; i++) {
            fullPath.Add(blockPath[i]);
        }

        fullPath.AddRange(changedPath);
        var pathText = string.Join(" > ", fullPath.Select(FormatPathSegment));
        if (isOriginal) {
            diagnostic.AffectedOriginalNodePath = pathText;
            diagnostic.AffectedOriginalNodeSpan = changedNode.SourceSpan;
        } else {
            diagnostic.AffectedFinalNodePath = pathText;
            diagnostic.AffectedFinalNodeSpan = changedNode.SourceSpan;
        }
    }

    private static ChangedNodePaths? FindChangedNodePaths(
        MarkdownSyntaxNode beforeNode,
        MarkdownSyntaxNode afterNode) {
        if (beforeNode == null || afterNode == null) {
            return null;
        }

        if (NodeFingerprint(beforeNode) == NodeFingerprint(afterNode)) {
            return null;
        }

        if (beforeNode.Kind != afterNode.Kind ||
            !string.Equals(beforeNode.CustomKind, afterNode.CustomKind, StringComparison.Ordinal)) {
            return new ChangedNodePaths(new[] { beforeNode }, new[] { afterNode });
        }

        var childChange = ComputeChildChange(beforeNode.Children, afterNode.Children);
        if (childChange.CountBefore == 1 && childChange.CountAfter == 1) {
            var beforeChild = beforeNode.Children[childChange.StartBefore];
            var afterChild = afterNode.Children[childChange.StartAfter];
            var childPaths = FindChangedNodePaths(beforeChild, afterChild);
            if (childPaths != null) {
                var originalPath = new List<MarkdownSyntaxNode>(childPaths.OriginalPath.Count + 1) {
                    beforeNode
                };
                originalPath.AddRange(childPaths.OriginalPath);

                var finalPath = new List<MarkdownSyntaxNode>(childPaths.FinalPath.Count + 1) {
                    afterNode
                };
                finalPath.AddRange(childPaths.FinalPath);
                return new ChangedNodePaths(originalPath, finalPath);
            }

            return new ChangedNodePaths(
                new[] { beforeNode, beforeChild },
                new[] { afterNode, afterChild });
        }

        return new ChangedNodePaths(new[] { beforeNode }, new[] { afterNode });
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

    private sealed class ChangedNodePaths {
        internal ChangedNodePaths(
            IReadOnlyList<MarkdownSyntaxNode> originalPath,
            IReadOnlyList<MarkdownSyntaxNode> finalPath) {
            OriginalPath = originalPath ?? Array.Empty<MarkdownSyntaxNode>();
            FinalPath = finalPath ?? Array.Empty<MarkdownSyntaxNode>();
        }

        internal IReadOnlyList<MarkdownSyntaxNode> OriginalPath { get; }

        internal IReadOnlyList<MarkdownSyntaxNode> FinalPath { get; }
    }
}
