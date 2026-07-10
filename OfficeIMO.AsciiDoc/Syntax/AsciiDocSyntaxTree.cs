namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Lossless AsciiDoc syntax tree rooted in the original source text.
/// </summary>
public sealed class AsciiDocSyntaxTree {
    internal AsciiDocSyntaxTree(AsciiDocSourceText source, AsciiDocSyntaxNode root) {
        Source = source ?? throw new ArgumentNullException(nameof(source));
        Root = root ?? throw new ArgumentNullException(nameof(root));
        IsLossless = ValidateNodeCoverage(source.Text, root);
    }

    /// <summary>Original source text and line mapping.</summary>
    public AsciiDocSourceText Source { get; }

    /// <summary>Document root node.</summary>
    public AsciiDocSyntaxNode Root { get; }

    /// <summary>
    /// True when root children are contiguous, cover the complete input, and retain the exact source characters.
    /// </summary>
    public bool IsLossless { get; }

    internal static bool ValidateNodeCoverage(string source, AsciiDocSyntaxNode node) {
        if (node.Span.Start.Offset < 0 || node.Span.End.Offset > source.Length) return false;
        if (!string.Equals(node.OriginalText, source.Substring(node.Span.Start.Offset, node.Span.Length), StringComparison.Ordinal)) return false;
        if (node.Children.Count == 0) return true;

        int expectedOffset = node.Span.Start.Offset;
        for (int index = 0; index < node.Children.Count; index++) {
            AsciiDocSyntaxNode child = node.Children[index];
            if (child.Span.Start.Offset != expectedOffset) return false;
            if (child.Span.End.Offset < child.Span.Start.Offset || child.Span.End.Offset > node.Span.End.Offset) return false;
            if (!ValidateNodeCoverage(source, child)) return false;
            expectedOffset = child.Span.End.Offset;
        }

        return expectedOffset == node.Span.End.Offset;
    }
}
