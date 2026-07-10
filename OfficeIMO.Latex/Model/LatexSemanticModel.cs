namespace OfficeIMO.Latex;

internal interface ILatexSourceEdit {
    bool IsModified { get; }
    LatexSourceSpan EditSpan { get; }
    string Replacement { get; }
}

/// <summary>Required or optional command argument.</summary>
public sealed class LatexArgument : ILatexSourceEdit {
    private string _content;
    private bool _isModified;

    internal LatexArgument(LatexSyntaxNode syntax, LatexSourceText source) {
        Syntax = syntax;
        IsOptional = syntax.Kind == LatexSyntaxKind.OptionalGroup;
        bool closed = syntax.Children.Count >= 2 && syntax.Children[syntax.Children.Count - 1].Kind == LatexSyntaxKind.GroupDelimiter;
        int start = syntax.Children.Count == 0 ? syntax.Span.Start.Offset : syntax.Children[0].Span.End.Offset;
        int end = closed ? syntax.Children[syntax.Children.Count - 1].Span.Start.Offset : syntax.Span.End.Offset;
        ContentSpan = source.CreateSpan(start, end);
        _content = source.Text.Substring(start, end - start);
        IsTerminated = closed;
    }

    /// <summary>Lossless group syntax.</summary>
    public LatexSyntaxNode Syntax { get; }
    /// <summary>True for square-bracket optional arguments.</summary>
    public bool IsOptional { get; }
    /// <summary>True when a closing delimiter was present.</summary>
    public bool IsTerminated { get; }
    /// <summary>Content span excluding delimiters.</summary>
    public LatexSourceSpan ContentSpan { get; }
    /// <summary>Argument content excluding delimiters.</summary>
    public string Content {
        get => _content;
        set {
            string normalized = value ?? string.Empty;
            if (string.Equals(_content, normalized, StringComparison.Ordinal)) return;
            _content = normalized;
            _isModified = true;
        }
    }
    /// <summary>True when content changed.</summary>
    public bool IsModified => _isModified;
    LatexSourceSpan ILatexSourceEdit.EditSpan => ContentSpan;
    string ILatexSourceEdit.Replacement => Content;
}

/// <summary>Source-backed LaTeX command without macro execution.</summary>
public sealed class LatexCommand {
    private readonly IReadOnlyList<LatexArgument> _arguments;

    internal LatexCommand(LatexSyntaxNode syntax, LatexSourceText source) {
        Syntax = syntax;
        Name = syntax.Value ?? string.Empty;
        _arguments = syntax.Children
            .Where(static child => child.Kind == LatexSyntaxKind.RequiredGroup || child.Kind == LatexSyntaxKind.OptionalGroup)
            .Select(child => new LatexArgument(child, source))
            .ToArray();
    }

    /// <summary>Lossless command syntax.</summary>
    public LatexSyntaxNode Syntax { get; }
    /// <summary>Command name without leading backslash.</summary>
    public string Name { get; }
    /// <summary>Bound optional and required arguments in source order.</summary>
    public IReadOnlyList<LatexArgument> Arguments => _arguments;
    /// <summary>True when the OfficeIMO profile assigns semantics to the command.</summary>
    public bool IsProfileKnown => LatexProfileCatalog.IsKnownCommand(Name);
    /// <summary>True when a recognized starred command modifier immediately follows the control word.</summary>
    public bool IsStarred => Syntax.Children.Any(static child =>
        child.Kind == LatexSyntaxKind.Text && string.Equals(child.OriginalText, "*", StringComparison.Ordinal));
    /// <summary>True when any argument changed.</summary>
    public bool IsModified => Arguments.Any(static argument => argument.IsModified);
    /// <summary>Returns the required argument at a zero-based required-only index.</summary>
    public LatexArgument? GetRequiredArgument(int index) => Arguments.Where(static argument => !argument.IsOptional).Skip(index).FirstOrDefault();
    /// <summary>Returns the optional argument at a zero-based optional-only index.</summary>
    public LatexArgument? GetOptionalArgument(int index) => Arguments.Where(static argument => argument.IsOptional).Skip(index).FirstOrDefault();
}

/// <summary>Source-backed begin/end LaTeX environment.</summary>
public sealed class LatexEnvironment : ILatexSourceEdit {
    private string _content;
    private bool _isModified;

    internal LatexEnvironment(
        LatexSyntaxNode syntax,
        LatexCommand beginCommand,
        LatexCommand? endCommand,
        LatexSourceText source) {
        Syntax = syntax;
        Name = syntax.Value ?? string.Empty;
        BeginCommand = beginCommand;
        EndCommand = endCommand;
        int start = beginCommand.Syntax.Span.End.Offset;
        int end = endCommand?.Syntax.Span.Start.Offset ?? syntax.Span.End.Offset;
        ContentSpan = source.CreateSpan(start, end);
        _content = source.Text.Substring(start, end - start);
    }

    /// <summary>Lossless environment syntax.</summary>
    public LatexSyntaxNode Syntax { get; }
    /// <summary>Environment name.</summary>
    public string Name { get; }
    /// <summary>Opening command.</summary>
    public LatexCommand BeginCommand { get; }
    /// <summary>Closing command, or null when unterminated.</summary>
    public LatexCommand? EndCommand { get; }
    /// <summary>True when a matching end command exists.</summary>
    public bool IsTerminated => EndCommand != null;
    /// <summary>Content span between begin and end commands.</summary>
    public LatexSourceSpan ContentSpan { get; }
    /// <summary>Exact environment body until edited.</summary>
    public string Content {
        get => _content;
        set {
            string normalized = value ?? string.Empty;
            if (string.Equals(_content, normalized, StringComparison.Ordinal)) return;
            _content = normalized;
            _isModified = true;
        }
    }
    /// <summary>True for a known display math environment.</summary>
    public bool IsMath => LatexProfileCatalog.IsMathEnvironment(Name);
    /// <summary>True for a known OfficeIMO profile environment.</summary>
    public bool IsProfileKnown => LatexProfileCatalog.IsKnownEnvironment(Name);
    /// <summary>True when body content changed.</summary>
    public bool IsModified => _isModified;
    LatexSourceSpan ILatexSourceEdit.EditSpan => ContentSpan;
    string ILatexSourceEdit.Replacement => Content;
}

/// <summary>LaTeX math delimiter kind.</summary>
public enum LatexMathKind {
    /// <summary>Single-dollar inline math.</summary>
    InlineDollar = 0,
    /// <summary>Parenthesized inline math.</summary>
    InlineParentheses,
    /// <summary>Double-dollar display math.</summary>
    DisplayDollar,
    /// <summary>Bracketed display math.</summary>
    DisplayBrackets,
    /// <summary>Named display math environment.</summary>
    Environment
}

/// <summary>Source-backed LaTeX math region; expressions are transported, not typeset.</summary>
public sealed class LatexMath : ILatexSourceEdit {
    private string _content;
    private bool _isModified;

    internal LatexMath(LatexSyntaxNode syntax, LatexSourceText source) {
        Syntax = syntax;
        Delimiter = syntax.Value ?? string.Empty;
        Kind = Delimiter == "$" ? LatexMathKind.InlineDollar
            : Delimiter == "$$" ? LatexMathKind.DisplayDollar
            : Delimiter == "\\(" ? LatexMathKind.InlineParentheses
            : LatexMathKind.DisplayBrackets;
        int start = syntax.Children.Count == 0 ? syntax.Span.Start.Offset : syntax.Children[0].Span.End.Offset;
        bool closed = syntax.Children.Count >= 2 && syntax.Children[syntax.Children.Count - 1].Kind == LatexSyntaxKind.MathDelimiter;
        int end = closed ? syntax.Children[syntax.Children.Count - 1].Span.Start.Offset : syntax.Span.End.Offset;
        ContentSpan = source.CreateSpan(start, end);
        _content = source.Text.Substring(start, end - start);
        IsTerminated = closed;
    }

    internal LatexMath(LatexEnvironment environment) {
        Environment = environment;
        Syntax = environment.Syntax;
        Delimiter = environment.Name;
        Kind = LatexMathKind.Environment;
        ContentSpan = environment.ContentSpan;
        _content = environment.Content;
        IsTerminated = environment.IsTerminated;
    }

    /// <summary>Lossless syntax.</summary>
    public LatexSyntaxNode Syntax { get; }
    /// <summary>Delimiter text or environment name.</summary>
    public string Delimiter { get; }
    /// <summary>Math kind.</summary>
    public LatexMathKind Kind { get; }
    /// <summary>True when closing syntax exists.</summary>
    public bool IsTerminated { get; }
    /// <summary>Expression span excluding delimiters.</summary>
    public LatexSourceSpan ContentSpan { get; }
    /// <summary>Original LaTeX expression source until edited.</summary>
    public string Content {
        get => Environment?.Content ?? _content;
        set {
            if (Environment != null) { Environment.Content = value; return; }
            string normalized = value ?? string.Empty;
            if (string.Equals(_content, normalized, StringComparison.Ordinal)) return;
            _content = normalized;
            _isModified = true;
        }
    }
    /// <summary>Backing math environment, when applicable.</summary>
    public LatexEnvironment? Environment { get; }
    /// <summary>True when expression changed.</summary>
    public bool IsModified => Environment?.IsModified ?? _isModified;
    LatexSourceSpan ILatexSourceEdit.EditSpan => ContentSpan;
    string ILatexSourceEdit.Replacement => Content;
}

/// <summary>Profile heading command.</summary>
public sealed class LatexHeading {
    internal LatexHeading(LatexCommand command, int level) {
        Command = command;
        Level = level;
    }

    /// <summary>Backing command.</summary>
    public LatexCommand Command { get; }
    /// <summary>Section level, where 1 is section.</summary>
    public int Level { get; }
    /// <summary>Heading title argument.</summary>
    public string Title {
        get => Command.GetRequiredArgument(0)?.Content ?? string.Empty;
        set {
            LatexArgument argument = Command.GetRequiredArgument(0) ?? throw new InvalidOperationException("Heading has no required title argument.");
            argument.Content = value ?? string.Empty;
        }
    }
    /// <summary>Optional short title.</summary>
    public string? ShortTitle {
        get => Command.GetOptionalArgument(0)?.Content;
        set {
            LatexArgument? argument = Command.GetOptionalArgument(0);
            if (argument == null) throw new InvalidOperationException("Heading has no optional short-title argument in source.");
            argument.Content = value ?? string.Empty;
        }
    }
}

/// <summary>Source-backed paragraph region within the document environment.</summary>
public sealed class LatexParagraph : ILatexSourceEdit {
    private string _content;
    private bool _isModified;

    internal LatexParagraph(LatexSourceSpan span, string content) {
        Span = span;
        _content = content;
    }

    /// <summary>Exact paragraph span excluding surrounding blank lines.</summary>
    public LatexSourceSpan Span { get; }
    /// <summary>Paragraph LaTeX source.</summary>
    public string Content {
        get => _content;
        set {
            string normalized = value ?? string.Empty;
            if (string.Equals(_content, normalized, StringComparison.Ordinal)) return;
            _content = normalized;
            _isModified = true;
        }
    }
    /// <summary>True when content changed.</summary>
    public bool IsModified => _isModified;
    LatexSourceSpan ILatexSourceEdit.EditSpan => Span;
    string ILatexSourceEdit.Replacement => Content;
}

internal static class LatexProfileCatalog {
    private static readonly HashSet<string> Commands = new HashSet<string>(StringComparer.Ordinal) {
        "documentclass", "usepackage", "title", "author", "date", "maketitle",
        "part", "chapter", "section", "subsection", "subsubsection", "paragraph", "subparagraph",
        "textbf", "textit", "emph", "texttt", "underline", "textsuperscript", "sout", "label", "ref", "pageref", "cite", "citep", "citet",
        "includegraphics", "caption", "item", "footnote", "url", "href", "begin", "end",
        "newcommand", "renewcommand", "providecommand", "newtheorem", "autoref", "eqref", "nocite",
        "bibliography", "bibliographystyle", "multicolumn", "multirow", "hline", "toprule", "midrule", "bottomrule"
    };

    private static readonly HashSet<string> Environments = new HashSet<string>(StringComparer.Ordinal) {
        "document", "itemize", "enumerate", "description", "figure", "table", "tabular", "quote", "quotation", "verbatim",
        "equation", "equation*", "align", "align*", "gather", "gather*", "multline", "multline*",
        "theorem", "lemma", "proposition", "corollary", "definition", "remark", "proof"
    };

    private static readonly HashSet<string> MathEnvironments = new HashSet<string>(StringComparer.Ordinal) {
        "equation", "equation*", "align", "align*", "gather", "gather*", "multline", "multline*"
    };

    internal static bool IsKnownCommand(string name) => Commands.Contains(name);
    internal static bool IsKnownEnvironment(string name) => Environments.Contains(name);
    internal static bool IsMathEnvironment(string name) => MathEnvironments.Contains(name);
}
