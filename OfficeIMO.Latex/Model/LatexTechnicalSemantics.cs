namespace OfficeIMO.Latex;

/// <summary>Itemize, enumerate, or description list kind.</summary>
public enum LatexListKind {
    /// <summary>Bulleted itemize.</summary>
    Unordered = 0,
    /// <summary>Numbered enumerate.</summary>
    Ordered,
    /// <summary>Term/description list.</summary>
    Description
}

/// <summary>Source-backed LaTeX list item.</summary>
public sealed class LatexListItem : ILatexSourceEdit {
    private string _content;
    private bool _isModified;

    internal LatexListItem(LatexCommand itemCommand, LatexSourceSpan contentSpan, string content) {
        ItemCommand = itemCommand;
        ContentSpan = contentSpan;
        _content = content;
    }

    /// <summary>Backing <c>\item</c> command.</summary>
    public LatexCommand ItemCommand { get; }
    /// <summary>Optional item label or description term.</summary>
    public string? Label {
        get => ItemCommand.GetOptionalArgument(0)?.Content;
        set {
            LatexArgument? argument = ItemCommand.GetOptionalArgument(0);
            if (argument == null) throw new InvalidOperationException("List item has no optional label in source.");
            argument.Content = value ?? string.Empty;
        }
    }
    /// <summary>Content span after <c>\item</c> through the next item or environment end.</summary>
    public LatexSourceSpan ContentSpan { get; }
    /// <summary>Item LaTeX source.</summary>
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

/// <summary>Semantic list environment.</summary>
public sealed class LatexList {
    internal LatexList(LatexEnvironment environment, LatexListKind kind, IReadOnlyList<LatexListItem> items) {
        Environment = environment;
        Kind = kind;
        Items = items;
    }

    /// <summary>Backing environment.</summary>
    public LatexEnvironment Environment { get; }
    /// <summary>List kind.</summary>
    public LatexListKind Kind { get; }
    /// <summary>Items in source order.</summary>
    public IReadOnlyList<LatexListItem> Items { get; }
}

/// <summary>Source-backed <c>\includegraphics</c> invocation.</summary>
public sealed class LatexImage {
    internal LatexImage(LatexCommand command) { Command = command; }
    /// <summary>Backing command.</summary>
    public LatexCommand Command { get; }
    /// <summary>Resource target.</summary>
    public string Target {
        get => Command.GetRequiredArgument(0)?.Content ?? string.Empty;
        set {
            LatexArgument argument = Command.GetRequiredArgument(0) ?? throw new InvalidOperationException("Image command has no target.");
            argument.Content = value ?? string.Empty;
        }
    }
    /// <summary>Raw optional graphics key/value list.</summary>
    public string? Options => Command.GetOptionalArgument(0)?.Content;
}

/// <summary>Semantic figure environment.</summary>
public sealed class LatexFigure {
    internal LatexFigure(
        LatexEnvironment environment,
        IReadOnlyList<LatexImage> images,
        LatexCommand? captionCommand,
        LatexCommand? labelCommand) {
        Environment = environment;
        Images = images;
        CaptionCommand = captionCommand;
        LabelCommand = labelCommand;
    }

    /// <summary>Backing figure environment.</summary>
    public LatexEnvironment Environment { get; }
    /// <summary>Included graphics.</summary>
    public IReadOnlyList<LatexImage> Images { get; }
    /// <summary>Caption command.</summary>
    public LatexCommand? CaptionCommand { get; }
    /// <summary>Label command.</summary>
    public LatexCommand? LabelCommand { get; }
    /// <summary>Caption text.</summary>
    public string? Caption {
        get => CaptionCommand?.GetRequiredArgument(0)?.Content;
        set {
            LatexArgument argument = CaptionCommand?.GetRequiredArgument(0) ?? throw new InvalidOperationException("Figure has no caption command in source.");
            argument.Content = value ?? string.Empty;
        }
    }
    /// <summary>Cross-reference label.</summary>
    public string? Label => LabelCommand?.GetRequiredArgument(0)?.Content;
}

/// <summary>Source-backed tabular cell.</summary>
public sealed class LatexTableCell : ILatexSourceEdit {
    private string _content;
    private bool _isModified;

    internal LatexTableCell(LatexSourceSpan span, string content, int rowIndex, int columnIndex) {
        Span = span;
        _content = content;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
    }

    /// <summary>Exact trimmed cell span.</summary>
    public LatexSourceSpan Span { get; }
    /// <summary>Zero-based row.</summary>
    public int RowIndex { get; }
    /// <summary>Zero-based column.</summary>
    public int ColumnIndex { get; }
    /// <summary>Cell LaTeX source.</summary>
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

/// <summary>Logical tabular row.</summary>
public sealed class LatexTableRow {
    internal LatexTableRow(int index, IReadOnlyList<LatexTableCell> cells) { Index = index; Cells = cells; }
    /// <summary>Zero-based row.</summary>
    public int Index { get; }
    /// <summary>Cells.</summary>
    public IReadOnlyList<LatexTableCell> Cells { get; }
}

/// <summary>Semantic tabular environment.</summary>
public sealed class LatexTable {
    internal LatexTable(LatexEnvironment environment, string columnSpecification, IReadOnlyList<LatexTableRow> rows) {
        Environment = environment;
        ColumnSpecification = columnSpecification;
        Rows = rows;
    }

    /// <summary>Backing tabular environment.</summary>
    public LatexEnvironment Environment { get; }
    /// <summary>Raw column specification.</summary>
    public string ColumnSpecification { get; }
    /// <summary>Rows and cells.</summary>
    public IReadOnlyList<LatexTableRow> Rows { get; }
}

/// <summary>Source-backed citation command.</summary>
public sealed class LatexCitation {
    internal LatexCitation(LatexCommand command, IReadOnlyList<string> keys) { Command = command; Keys = keys; }
    /// <summary>Backing cite command.</summary>
    public LatexCommand Command { get; }
    /// <summary>Citation keys.</summary>
    public IReadOnlyList<string> Keys { get; }
    /// <summary>Optional prenote.</summary>
    public string? Prenote => Command.GetOptionalArgument(0)?.Content;
}

/// <summary>Source-backed reference command.</summary>
public sealed class LatexReference {
    internal LatexReference(LatexCommand command, string target) { Command = command; Target = target; }
    /// <summary>Backing command.</summary>
    public LatexCommand Command { get; }
    /// <summary>Reference label target.</summary>
    public string Target { get; }
}

/// <summary>Source-backed label declaration.</summary>
public sealed class LatexLabel {
    internal LatexLabel(LatexCommand command, string name) { Command = command; Name = name; }
    /// <summary>Backing command.</summary>
    public LatexCommand Command { get; }
    /// <summary>Label name.</summary>
    public string Name { get; }
}

/// <summary>Theorem-like environment with optional title and label.</summary>
public sealed class LatexTheorem {
    internal LatexTheorem(LatexEnvironment environment, LatexCommand? labelCommand) {
        Environment = environment;
        LabelCommand = labelCommand;
    }
    /// <summary>Backing theorem environment.</summary>
    public LatexEnvironment Environment { get; }
    /// <summary>Theorem kind/environment name.</summary>
    public string Kind => Environment.Name;
    /// <summary>Optional title on the begin command.</summary>
    public string? Title => Environment.BeginCommand.GetOptionalArgument(0)?.Content;
    /// <summary>Body source.</summary>
    public string Content => Environment.Content;
    /// <summary>Label command.</summary>
    public LatexCommand? LabelCommand { get; }
    /// <summary>Label.</summary>
    public string? Label => LabelCommand?.GetRequiredArgument(0)?.Content;
}
