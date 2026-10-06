namespace OfficeIMO.Latex.Markdown;

// Only reached inline notes produce definitions. Notes in opaque command arguments
// remain source fallbacks, and repeated scalar reads reuse the same target label.
internal sealed class LatexFootnoteProjection {
    private readonly Dictionary<LatexSyntaxNode, LatexFootnoteUse> _bySyntax = new();
    private readonly List<LatexFootnoteUse> _used = new();
    internal IReadOnlyList<LatexFootnoteUse> Used => _used;

    internal string Register(LatexCommand command) {
        if (_bySyntax.TryGetValue(command.Syntax, out LatexFootnoteUse? existing)) return existing.Label;
        string label = "latex-" + (_used.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var note = new LatexFootnoteUse(command, label);
        _bySyntax.Add(command.Syntax, note);
        _used.Add(note);
        return label;
    }
}

internal sealed class LatexFootnoteUse {
    internal LatexFootnoteUse(LatexCommand command, string label) {
        Command = command;
        Label = label;
        Body = command.GetRequiredArgument(0)!;
    }
    internal LatexCommand Command { get; }
    internal string Label { get; }
    internal LatexArgument Body { get; }
}
