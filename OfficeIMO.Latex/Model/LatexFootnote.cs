namespace OfficeIMO.Latex;

/// <summary>A source-backed footnote. The optional mark is retained without evaluating TeX counters.</summary>
public sealed class LatexFootnote {
    internal LatexFootnote(LatexCommand command) { Command = command; }

    /// <summary>The backing <c>\footnote</c> command and its original syntax.</summary>
    public LatexCommand Command { get; }

    /// <summary>The body argument, including its source span and original braced or single-token form.</summary>
    public LatexArgument Body => Command.GetRequiredArgument(0)!;

    /// <summary>The footnote body as LaTeX source. Changes participate in normal conflict-checked source writing.</summary>
    public string Content {
        get => Body.Content;
        set => Body.Content = value;
    }

    /// <summary>The original optional mark; no automatic footnote number is calculated.</summary>
    public string? Mark {
        get => Command.GetOptionalArgument(0)?.Content;
        set {
            LatexArgument argument = Command.GetOptionalArgument(0)
                ?? throw new InvalidOperationException("Footnote has no optional mark in source.");
            argument.Content = value ?? string.Empty;
        }
    }
}
