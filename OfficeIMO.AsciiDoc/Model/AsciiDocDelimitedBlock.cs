namespace OfficeIMO.AsciiDoc;

/// <summary>Common AsciiDoc delimited block kind.</summary>
public enum AsciiDocDelimitedBlockKind {
    /// <summary>Listing or source-code block.</summary>
    Listing = 0,
    /// <summary>Literal block.</summary>
    Literal,
    /// <summary>Example block.</summary>
    Example,
    /// <summary>Sidebar block.</summary>
    Sidebar,
    /// <summary>Quote block.</summary>
    Quote,
    /// <summary>Passthrough block.</summary>
    Passthrough,
    /// <summary>Open block.</summary>
    Open,
    /// <summary>Table block.</summary>
    Table,
    /// <summary>Comment block.</summary>
    Comment
}

/// <summary>Source-preserving common AsciiDoc delimited block.</summary>
public class AsciiDocDelimitedBlock : AsciiDocBlock {
    private string _content;
    private bool _contentWasAssigned;

    internal AsciiDocDelimitedBlock(
        AsciiDocSyntaxNode syntax,
        AsciiDocDelimitedBlockKind kind,
        string delimiter,
        string openingText,
        string content,
        string closingText,
        bool isTerminated,
        string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        Kind = kind;
        Delimiter = delimiter;
        OpeningText = openingText;
        _content = content;
        ClosingText = closingText;
        IsTerminated = isTerminated;
    }

    /// <summary>Delimited block kind.</summary>
    public AsciiDocDelimitedBlockKind Kind { get; }

    /// <summary>Opening and closing delimiter text without a line ending.</summary>
    public string Delimiter { get; }

    /// <summary>Exact original opening delimiter line.</summary>
    public string OpeningText { get; }

    /// <summary>Exact original closing delimiter line, or empty when unterminated.</summary>
    public string ClosingText { get; }

    /// <summary>True when a matching closing delimiter was present.</summary>
    public bool IsTerminated { get; }

    /// <summary>Content between delimiters, including its original internal line endings until edited.</summary>
    public string Content {
        get => _content;
        set {
            if (SetValue(ref _content, value ?? string.Empty)) _contentWasAssigned = true;
        }
    }

    /// <summary>True when <see cref="Content"/> was replaced directly.</summary>
    protected bool IsContentAssigned => _contentWasAssigned;

    /// <summary>Admonition kind for a styled compound block, when applicable.</summary>
    public AsciiDocAdmonitionKind? AdmonitionKind {
        get {
            switch (Style?.ToUpperInvariant()) {
                case "NOTE": return AsciiDocAdmonitionKind.Note;
                case "TIP": return AsciiDocAdmonitionKind.Tip;
                case "IMPORTANT": return AsciiDocAdmonitionKind.Important;
                case "WARNING": return AsciiDocAdmonitionKind.Warning;
                case "CAUTION": return AsciiDocAdmonitionKind.Caution;
                default: return null;
            }
        }
    }

    /// <summary>True when this is a display STEM block.</summary>
    public bool IsStem => string.Equals(Style, "stem", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Style, "latexmath", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Style, "asciimath", StringComparison.OrdinalIgnoreCase);

    internal override string WriteCore(AsciiDocWriterContext context) {
        if (context.Mode == AsciiDocWriterMode.Preserve) {
            string content = Content;
            if (IsTerminated && content.Length > 0 && !AsciiDocText.EndsWithLineEnding(content)) {
                content += context.LineEnding;
            }
            return OpeningText + content + ClosingText;
        }

        var builder = new StringBuilder();
        builder.Append(Delimiter).Append(context.LineEnding);
        string normalized = AsciiDocText.NormalizeLineEndings(Content, context.LineEnding);
        builder.Append(normalized);
        if (IsTerminated) {
            if (normalized.Length > 0 && !AsciiDocText.EndsWithLineEnding(normalized)) builder.Append(context.LineEnding);
            builder.Append(Delimiter).Append(EffectiveTrailingLineEnding(context));
        }
        return builder.ToString();
    }
}
