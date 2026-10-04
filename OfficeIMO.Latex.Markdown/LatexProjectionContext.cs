using System.Threading;

namespace OfficeIMO.Latex.Markdown;

// Indexes belong to one current source snapshot. Block and recursive inline projection
// must query their source range instead of rescanning the complete document.
internal sealed class LatexProjectionContext {
    private const long MaximumProjectedTableCells = 65536;
    private long _projectedTableCells;
    private readonly LatexInlineCandidate[] _inlines;
    private readonly LatexProjectionComment[] _comments;
    private readonly LatexLabel[] _labels;
    private readonly Dictionary<string, LatexCommand> _firstCommands = new(StringComparer.Ordinal);
    private readonly Dictionary<LatexSyntaxNode, Dictionary<string, LatexCommand>> _directCommands = new();
    private readonly Dictionary<LatexSyntaxNode, LatexCommand> _theoremLabels = new();
    private readonly Dictionary<LatexSyntaxNode, LatexEnvironment> _environments = new();
    private readonly HashSet<LatexEnvironment> _semanticEnvironments = new();
    private readonly IReadOnlyList<LatexSourceSpan> _excluded = Array.Empty<LatexSourceSpan>();

    internal LatexProjectionContext(LatexDocument document, CancellationToken cancellationToken) {
        Document = document;
        CancellationToken = cancellationToken;
        var inlines = new List<LatexInlineCandidate>();
        var comments = new List<LatexProjectionComment>();
        foreach (LatexEnvironment environment in document.Environments) {
            CheckCancellation();
            _environments[environment.Syntax] = environment;
        }
        foreach (LatexCommand command in document.Commands) {
            CheckCancellation();
            if (!LatexSemanticBuilder.IsActiveSyntax(command.Syntax)) continue;
            inlines.Add(new LatexInlineCandidate(command.Syntax.Span, command, null, null));
            if (command.Syntax.Span.End.Offset > (document.Body?.Syntax.Span.End.Offset ?? document.Source.Text.Length)) continue;
            LatexSyntaxNode? owner = command.Syntax.Parent;
            while (owner != null && owner.Kind != LatexSyntaxKind.Environment) owner = owner.Parent;
            if (owner != null && command.Name == "label" && !IsInsidePreservedArgument(command.Syntax) && !_theoremLabels.ContainsKey(owner))
                _theoremLabels.Add(owner, command);
            if (LatexSemanticBuilder.IsInsideCommandArgument(command.Syntax)) continue;
            if ((owner == null || ReferenceEquals(owner, document.Body?.Syntax)) && !_firstCommands.ContainsKey(command.Name))
                _firstCommands.Add(command.Name, command);
            if (owner == null) continue;
            if (!_directCommands.TryGetValue(owner, out var commands)) {
                commands = new Dictionary<string, LatexCommand>(StringComparer.Ordinal);
                _directCommands.Add(owner, commands);
            }
            if (!commands.ContainsKey(command.Name)) commands.Add(command.Name, command);
        }
        foreach (LatexMath math in document.Math) {
            CheckCancellation();
            if (math.Kind != LatexMathKind.Environment) inlines.Add(new LatexInlineCandidate(math.Syntax.Span, null, math, null));
            if (math.Environment != null) _semanticEnvironments.Add(math.Environment);
        }
        var verbatim = new List<LatexSyntaxNode>();
        foreach (LatexSyntaxNode node in document.SyntaxTree.Root.DescendantsAndSelf()) {
            CheckCancellation();
            if (node.Kind != LatexSyntaxKind.Verbatim) continue;
            if (node.Value == "comment") comments.Add(new LatexProjectionComment(node.Span, node));
            if (!LatexSemanticBuilder.IsActiveSyntax(node)) continue;
            verbatim.Add(node);
            inlines.Add(new LatexInlineCandidate(node.Span, null, null, node));
        }
        Verbatim = verbatim;
        foreach (LatexToken token in document.Tokens) {
            CheckCancellation();
            if (token.Kind != LatexTokenKind.Comment) continue;
            int end = token.Span.End.Offset;
            string source = document.Source.Text;
            if (end < source.Length && source[end] == '\r') end++;
            if (end < source.Length && source[end] == '\n') end++;
            comments.Add(new LatexProjectionComment(document.Source.CreateSpan(token.Span.Start.Offset, end), null));
        }
        foreach (LatexList item in document.Lists) { CheckCancellation(); _semanticEnvironments.Add(item.Environment); }
        foreach (LatexFigure item in document.Figures) { CheckCancellation(); _semanticEnvironments.Add(item.Environment); }
        foreach (LatexTheorem item in document.Theorems) { CheckCancellation(); _semanticEnvironments.Add(item.Environment); }
        foreach (LatexTable item in document.Tables) {
            CheckCancellation();
            _semanticEnvironments.Add(item.Environment);
            LatexEnvironment? container = FindAncestorEnvironment(item.Environment, "table");
            if (container != null) _semanticEnvironments.Add(container);
        }
        HasPart = document.Body != null && FindDirectCommand(document.Body, "part") != null;
        _inlines = inlines.OrderBy(static item => item.Span.Start.Offset).ThenByDescending(static item => item.Span.End.Offset).ToArray();
        _comments = comments.OrderBy(static item => item.Span.Start.Offset).ToArray();
        _labels = document.Labels.OrderBy(static item => item.Command.Syntax.Span.Start.Offset).ToArray();
        CheckCancellation();
    }

    internal LatexDocument Document { get; }
    internal CancellationToken CancellationToken { get; }
    internal IReadOnlyList<LatexSyntaxNode> Verbatim { get; }
    internal bool HasPart { get; }
    private LatexProjectionContext(LatexProjectionContext source, IReadOnlyList<LatexSourceSpan> excluded) {
        Document = source.Document;
        CancellationToken = source.CancellationToken;
        Verbatim = source.Verbatim;
        HasPart = source.HasPart;
        _inlines = source._inlines;
        _comments = source._comments;
        _labels = source._labels;
        _firstCommands = source._firstCommands;
        _directCommands = source._directCommands;
        _theoremLabels = source._theoremLabels;
        _environments = source._environments;
        _semanticEnvironments = source._semanticEnvironments;
        _excluded = excluded;
    }

    internal LatexProjectionContext WithExcludedSpans(IEnumerable<LatexSourceSpan> spans) {
        var excluded = new List<LatexSourceSpan>(_excluded);
        foreach (LatexSourceSpan span in spans) { CheckCancellation(); excluded.Add(span); }
        return new LatexProjectionContext(this, excluded);
    }

    internal bool IsExcluded(LatexSourceSpan span) => _excluded.Any(excluded =>
        excluded.Start.Offset <= span.Start.Offset && excluded.End.Offset >= span.End.Offset);
    internal void CheckCancellation() => CancellationToken.ThrowIfCancellationRequested();
    internal void ChargeProjectedTableCells(long cells) {
        if (cells < 0 || cells > MaximumProjectedTableCells - _projectedTableCells)
            throw new System.IO.InvalidDataException("LaTeX tables exceed the Markdown projection cell limit.");
        _projectedTableCells += cells;
    }
    internal LatexCommand? FirstCommand(string name) => _firstCommands.TryGetValue(name, out var command) ? command : null;
    internal string? DocumentClassName => FirstCommand("documentclass")?.GetRequiredArgument(0)?.Content.Trim();
    internal bool HasSemanticProjection(LatexEnvironment environment) => _semanticEnvironments.Contains(environment);
    internal LatexCommand? FindDirectCommand(LatexEnvironment? environment, string name) => environment != null
        && _directCommands.TryGetValue(environment.Syntax, out var commands) && commands.TryGetValue(name, out var command) ? command : null;
    internal LatexCommand? FindTheoremLabel(LatexEnvironment environment) =>
        _theoremLabels.TryGetValue(environment.Syntax, out var command) ? command : null;

    // Inline formatting is projected recursively, so a label can be removed from its
    // children without slicing the enclosing syntax. Preserved command arguments are
    // opaque source: their labels must stay inside that source rather than escape it.
    private static bool IsInsidePreservedArgument(LatexSyntaxNode node) {
        for (LatexSyntaxNode? parent = node.Parent; parent != null; parent = parent.Parent) {
            if (parent.Kind != LatexSyntaxKind.Command) continue;
            switch (parent.Value) {
                case "textbf": case "textit": case "emph": case "texttt": case "underline":
                case "textsuperscript": case "textsubscript": case "sout":
                    break;
                default:
                    return true;
            }
        }
        return false;
    }

    internal LatexEnvironment? FindAncestorEnvironment(LatexEnvironment source, string name) {
        LatexSyntaxNode? current = source.Syntax.Parent;
        while (current != null) {
            CheckCancellation();
            if (current.Kind == LatexSyntaxKind.Environment && current.Value == name)
                return _environments.TryGetValue(current, out var environment) ? environment : null;
            current = current.Parent;
        }
        return null;
    }

    internal LatexLabel? FindAdjacentLabel(LatexSourceSpan owner) {
        int index = LowerBound(_labels.Length, owner.End.Offset, i => _labels[i].Command.Syntax.Span.Start.Offset);
        if (index == _labels.Length) return null;
        LatexLabel label = _labels[index];
        for (int cursor = owner.End.Offset; cursor < label.Command.Syntax.Span.Start.Offset; cursor++) {
            if ((cursor & 1023) == 0) CheckCancellation();
            if (!char.IsWhiteSpace(Document.Source.Text[cursor])) return null;
        }
        CheckCancellation();
        return label;
    }

    internal IEnumerable<LatexInlineCandidate> InlineCandidates(int start, int end) {
        CheckCancellation();
        int first = LowerBound(_inlines.Length, start, i => _inlines[i].Span.Start.Offset);
        for (int index = first; index < _inlines.Length && _inlines[index].Span.Start.Offset < end; index++) {
            CheckCancellation();
            if (_inlines[index].Span.End.Offset <= end) yield return _inlines[index];
        }
    }

    internal IEnumerable<LatexProjectionComment> Comments(LatexSourceSpan span) {
        CheckCancellation();
        int first = LowerBound(_comments.Length, span.Start.Offset, i => _comments[i].Span.Start.Offset);
        // Lexical comments and opaque comment environments do not overlap. The preceding
        // comment may consume a line ending or straddle the requested content boundary.
        if (first > 0 && _comments[first - 1].Span.End.Offset > span.Start.Offset) first--;
        for (int index = first; index < _comments.Length && _comments[index].Span.Start.Offset < span.End.Offset; index++) {
            CheckCancellation();
            yield return _comments[index];
        }
    }

    internal string ExtractResidual(LatexSourceSpan span, IEnumerable<LatexSourceSpan> representedSpans) {
        var represented = new List<LatexSourceSpan>();
        foreach (LatexSourceSpan item in representedSpans.Concat(_excluded)) {
            CheckCancellation();
            if (item.End.Offset > span.Start.Offset && item.Start.Offset < span.End.Offset) represented.Add(item);
        }
        represented.Sort(static (left, right) => left.Start.Offset.CompareTo(right.Start.Offset));
        CheckCancellation();
        var output = new StringBuilder();
        int cursor = span.Start.Offset;
        foreach (LatexSourceSpan item in represented) {
            CheckCancellation();
            int start = Math.Max(cursor, item.Start.Offset);
            int end = Math.Min(span.End.Offset, item.End.Offset);
            AppendSource(output, cursor, start);
            cursor = Math.Max(cursor, end);
        }
        AppendSource(output, cursor, span.End.Offset);
        string result = output.ToString();
        CheckCancellation();
        return result;
    }

    private void AppendSource(StringBuilder output, int start, int end) {
        while (start < end) {
            CheckCancellation();
            int length = Math.Min(4096, end - start);
            output.Append(Document.Source.Text, start, length);
            start += length;
        }
    }

    private static int LowerBound(int length, int offset, Func<int, int> getStart) {
        int low = 0, high = length;
        while (low < high) {
            int middle = low + (high - low) / 2;
            if (getStart(middle) < offset) low = middle + 1;
            else high = middle;
        }
        return low;
    }
}

internal sealed class LatexInlineCandidate {
    internal LatexInlineCandidate(LatexSourceSpan span, LatexCommand? command, LatexMath? math, LatexSyntaxNode? verbatim) {
        Span = span; Command = command; Math = math; Verbatim = verbatim;
    }
    internal LatexSourceSpan Span { get; }
    internal LatexCommand? Command { get; }
    internal LatexMath? Math { get; }
    internal LatexSyntaxNode? Verbatim { get; }
}

internal sealed class LatexProjectionComment {
    internal LatexProjectionComment(LatexSourceSpan span, LatexSyntaxNode? opaqueSyntax) { Span = span; OpaqueSyntax = opaqueSyntax; }
    internal LatexSourceSpan Span { get; }
    internal LatexSyntaxNode? OpaqueSyntax { get; }
}
