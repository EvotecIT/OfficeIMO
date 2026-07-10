namespace OfficeIMO.AsciiDoc;

internal sealed class AsciiDocSyntaxFactory {
    private readonly AsciiDocSourceText _source;

    internal AsciiDocSyntaxFactory(AsciiDocSourceText source) {
        _source = source;
    }

    internal AsciiDocSourceText Source => _source;

    internal AsciiDocSyntaxNode Node(AsciiDocSyntaxKind kind, int start, int end, IReadOnlyList<AsciiDocSyntaxNode>? children = null) =>
        new AsciiDocSyntaxNode(
            kind,
            _source.CreateSpan(start, end),
            _source.Text.Substring(start, end - start),
            CompleteCoverage(start, end, children));

    internal void AddLineEnding(List<AsciiDocSyntaxNode> children, AsciiDocSourceLine line) {
        if (line.LineEndingLength > 0) children.Add(Node(AsciiDocSyntaxKind.LineEnding, line.ContentEnd, line.End));
    }

    private IReadOnlyList<AsciiDocSyntaxNode>? CompleteCoverage(
        int start,
        int end,
        IReadOnlyList<AsciiDocSyntaxNode>? children) {
        if (children == null || children.Count == 0) return children;

        var completed = new List<AsciiDocSyntaxNode>(children.Count + 2);
        int expected = start;
        for (int index = 0; index < children.Count; index++) {
            AsciiDocSyntaxNode child = children[index];
            if (child.Span.Start.Offset > expected) completed.Add(CreateTrivia(expected, child.Span.Start.Offset));
            completed.Add(child);
            expected = Math.Max(expected, child.Span.End.Offset);
        }
        if (expected < end) completed.Add(CreateTrivia(expected, end));
        return completed;
    }

    private AsciiDocSyntaxNode CreateTrivia(int start, int end) =>
        new AsciiDocSyntaxNode(
            AsciiDocSyntaxKind.Trivia,
            _source.CreateSpan(start, end),
            _source.Text.Substring(start, end - start));
}
