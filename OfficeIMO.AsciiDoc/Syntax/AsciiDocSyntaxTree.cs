namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Lossless AsciiDoc syntax tree rooted in the original source text.
/// </summary>
public sealed class AsciiDocSyntaxTree {
    internal AsciiDocSyntaxTree(AsciiDocSourceText source, AsciiDocSyntaxNode root) {
        Source = source ?? throw new ArgumentNullException(nameof(source));
        Root = root ?? throw new ArgumentNullException(nameof(root));
        IsLossless = ValidateNodeCoverage(source, root);
    }

    /// <summary>Original source text and line mapping.</summary>
    public AsciiDocSourceText Source { get; }

    /// <summary>Document root node.</summary>
    public AsciiDocSyntaxNode Root { get; }

    /// <summary>
    /// True when root children are contiguous, cover the complete input, and retain the exact source characters.
    /// </summary>
    public bool IsLossless { get; }

    internal static bool ValidateNodeCoverage(AsciiDocSourceText source, AsciiDocSyntaxNode node) {
        if (node.StartOffset < 0 || node.EndOffset > source.Text.Length) return false;
        if (!node.HasSource(source)) return false;
        if (node.Children.Count == 0) return true;

        int expectedOffset = node.StartOffset;
        for (int index = 0; index < node.Children.Count; index++) {
            AsciiDocSyntaxNode child = node.Children[index];
            if (child.StartOffset != expectedOffset) return false;
            if (child.EndOffset < child.StartOffset || child.EndOffset > node.EndOffset) return false;
            if (!ValidateNodeCoverage(source, child)) return false;
            expectedOffset = child.EndOffset;
        }

        return expectedOffset == node.EndOffset;
    }
}
