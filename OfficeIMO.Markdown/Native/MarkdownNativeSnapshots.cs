namespace OfficeIMO.Markdown;

/// <summary>
/// UI-safe snapshot of a native markdown document without parser object references.
/// </summary>
public sealed class MarkdownNativeDocumentSnapshot {
    internal MarkdownNativeDocumentSnapshot(
        MarkdownNativeDocumentSourceKind sourceKind,
        IReadOnlyList<MarkdownNativeBlockSnapshot> blocks,
        IReadOnlyList<MarkdownNativeDiagnosticSnapshot> diagnostics) {
        SourceKind = sourceKind;
        Blocks = blocks ?? Array.Empty<MarkdownNativeBlockSnapshot>();
        Diagnostics = diagnostics ?? Array.Empty<MarkdownNativeDiagnosticSnapshot>();
    }

    /// <summary>Identifies the markdown source backing this snapshot.</summary>
    public MarkdownNativeDocumentSourceKind SourceKind { get; }

    /// <summary>Top-level block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockSnapshot> Blocks { get; }

    /// <summary>Projection diagnostics.</summary>
    public IReadOnlyList<MarkdownNativeDiagnosticSnapshot> Diagnostics { get; }
}

/// <summary>
/// UI-safe snapshot of a native block.
/// </summary>
public sealed class MarkdownNativeBlockSnapshot {
    internal MarkdownNativeBlockSnapshot() {
        Fields = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        Inlines = Array.Empty<MarkdownNativeInlineSnapshot>();
        Children = Array.Empty<MarkdownNativeBlockSnapshot>();
        Items = Array.Empty<MarkdownNativeListItemSnapshot>();
        HeaderCells = Array.Empty<MarkdownNativeTableCellSnapshot>();
        Rows = Array.Empty<IReadOnlyList<MarkdownNativeTableCellSnapshot>>();
    }

    /// <summary>Stable block id.</summary>
    public string Id { get; internal set; } = string.Empty;

    /// <summary>Native block kind.</summary>
    public MarkdownNativeBlockKind Kind { get; internal set; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; internal set; }

    /// <summary>Common text payload when the block exposes one.</summary>
    public string? Text { get; internal set; }

    /// <summary>Common markdown payload when the block exposes one.</summary>
    public string? Markdown { get; internal set; }

    /// <summary>String fields for block-specific metadata.</summary>
    public IReadOnlyDictionary<string, string?> Fields { get; internal set; }

    /// <summary>Inline snapshots owned directly by this block.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Inlines { get; internal set; }

    /// <summary>Nested child block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockSnapshot> Children { get; internal set; }

    /// <summary>List item snapshots for native list blocks.</summary>
    public IReadOnlyList<MarkdownNativeListItemSnapshot> Items { get; internal set; }

    /// <summary>Table header cell snapshots.</summary>
    public IReadOnlyList<MarkdownNativeTableCellSnapshot> HeaderCells { get; internal set; }

    /// <summary>Table body row snapshots.</summary>
    public IReadOnlyList<IReadOnlyList<MarkdownNativeTableCellSnapshot>> Rows { get; internal set; }
}

/// <summary>
/// UI-safe snapshot of a native inline.
/// </summary>
public sealed class MarkdownNativeInlineSnapshot {
    internal MarkdownNativeInlineSnapshot(
        string id,
        MarkdownNativeInlineKind kind,
        MarkdownSyntaxKind syntaxKind,
        string text,
        string markdown,
        string literal,
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        IReadOnlyDictionary<string, string> metadata,
        IReadOnlyList<MarkdownNativeInlineSnapshot> children) {
        Id = id ?? string.Empty;
        Kind = kind;
        SyntaxKind = syntaxKind;
        Text = text ?? string.Empty;
        Markdown = markdown ?? string.Empty;
        Literal = literal ?? string.Empty;
        SourceSpan = sourceSpan;
        Metadata = metadata ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        Children = children ?? Array.Empty<MarkdownNativeInlineSnapshot>();
    }

    /// <summary>Stable inline id.</summary>
    public string Id { get; }

    /// <summary>Native inline kind.</summary>
    public MarkdownNativeInlineKind Kind { get; }

    /// <summary>Syntax kind that produced this inline.</summary>
    public MarkdownSyntaxKind SyntaxKind { get; }

    /// <summary>Plain text represented by this inline.</summary>
    public string Text { get; }

    /// <summary>Markdown represented by this inline.</summary>
    public string Markdown { get; }

    /// <summary>Literal syntax payload.</summary>
    public string Literal { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Metadata values such as target/title/source/alt.</summary>
    public IReadOnlyDictionary<string, string> Metadata { get; }

    /// <summary>Nested inline snapshots.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Children { get; }
}

/// <summary>
/// UI-safe snapshot of a native list item.
/// </summary>
public sealed class MarkdownNativeListItemSnapshot {
    internal MarkdownNativeListItemSnapshot(
        string id,
        string text,
        bool isTask,
        bool isChecked,
        int level,
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        IReadOnlyList<MarkdownNativeInlineSnapshot> inlines,
        IReadOnlyList<MarkdownNativeBlockSnapshot> children) {
        Id = id ?? string.Empty;
        Text = text ?? string.Empty;
        IsTask = isTask;
        IsChecked = isChecked;
        Level = level;
        SourceSpan = sourceSpan;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineSnapshot>();
        Children = children ?? Array.Empty<MarkdownNativeBlockSnapshot>();
    }

    /// <summary>Stable list item id.</summary>
    public string Id { get; }

    /// <summary>Plain text lead content.</summary>
    public string Text { get; }

    /// <summary>Whether the item is a task item.</summary>
    public bool IsTask { get; }

    /// <summary>Whether the task item is checked.</summary>
    public bool IsChecked { get; }

    /// <summary>Indentation level from the source item.</summary>
    public int Level { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Lead inline snapshots.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Inlines { get; }

    /// <summary>Nested child block snapshots.</summary>
    public IReadOnlyList<MarkdownNativeBlockSnapshot> Children { get; }
}

/// <summary>
/// UI-safe snapshot of a native table cell.
/// </summary>
public sealed class MarkdownNativeTableCellSnapshot {
    internal MarkdownNativeTableCellSnapshot(
        string text,
        string markdown,
        bool isHeader,
        int rowIndex,
        int columnIndex,
        ColumnAlignment alignment,
        MarkdownNativeSourceSpanSnapshot? sourceSpan,
        IReadOnlyList<MarkdownNativeInlineSnapshot> inlines) {
        Text = text ?? string.Empty;
        Markdown = markdown ?? string.Empty;
        IsHeader = isHeader;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Alignment = alignment;
        SourceSpan = sourceSpan;
        Inlines = inlines ?? Array.Empty<MarkdownNativeInlineSnapshot>();
    }

    /// <summary>Plain text cell content.</summary>
    public string Text { get; }

    /// <summary>Markdown cell content.</summary>
    public string Markdown { get; }

    /// <summary>Whether this is a header cell.</summary>
    public bool IsHeader { get; }

    /// <summary>Zero-based row index, or -1 for headers.</summary>
    public int RowIndex { get; }

    /// <summary>Zero-based column index.</summary>
    public int ColumnIndex { get; }

    /// <summary>Projected alignment.</summary>
    public ColumnAlignment Alignment { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Inline snapshots for cell content when available.</summary>
    public IReadOnlyList<MarkdownNativeInlineSnapshot> Inlines { get; }
}

/// <summary>
/// UI-safe source span snapshot.
/// </summary>
public sealed class MarkdownNativeSourceSpanSnapshot {
    internal MarkdownNativeSourceSpanSnapshot(MarkdownSourceSpan span) {
        StartLine = span.StartLine;
        StartColumn = span.StartColumn;
        EndLine = span.EndLine;
        EndColumn = span.EndColumn;
        StartOffset = span.StartOffset;
        EndOffset = span.EndOffset;
        Display = span.ToString();
    }

    /// <summary>1-based start line.</summary>
    public int StartLine { get; }

    /// <summary>1-based start column when available.</summary>
    public int? StartColumn { get; }

    /// <summary>1-based end line.</summary>
    public int EndLine { get; }

    /// <summary>1-based end column when available.</summary>
    public int? EndColumn { get; }

    /// <summary>0-based start offset when available.</summary>
    public int? StartOffset { get; }

    /// <summary>0-based end offset when available.</summary>
    public int? EndOffset { get; }

    /// <summary>Human-readable span display.</summary>
    public string Display { get; }
}

/// <summary>
/// UI-safe diagnostic snapshot.
/// </summary>
public sealed class MarkdownNativeDiagnosticSnapshot {
    internal MarkdownNativeDiagnosticSnapshot(MarkdownNativeDiagnostic diagnostic) {
        Id = diagnostic.Id;
        Message = diagnostic.Message;
        Severity = diagnostic.Severity;
        SourceSpan = diagnostic.SourceSpan.HasValue ? new MarkdownNativeSourceSpanSnapshot(diagnostic.SourceSpan.Value) : null;
        BlockId = diagnostic.Block?.Id;
    }

    /// <summary>Diagnostic id.</summary>
    public string Id { get; }

    /// <summary>Diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Diagnostic severity.</summary>
    public MarkdownNativeDiagnosticSeverity Severity { get; }

    /// <summary>Source span snapshot when available.</summary>
    public MarkdownNativeSourceSpanSnapshot? SourceSpan { get; }

    /// <summary>Associated block id when available.</summary>
    public string? BlockId { get; }
}
