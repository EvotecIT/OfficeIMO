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

    private static string FormatPathSegment(MarkdownSyntaxNode node) {
        if (string.IsNullOrWhiteSpace(node.CustomKind)) {
            return node.Kind.ToString();
        }

        return node.Kind + "(" + node.CustomKind + ")";
    }
}
