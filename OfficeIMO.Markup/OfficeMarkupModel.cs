using System.Collections.ObjectModel;

namespace OfficeIMO.Markup;

/// <summary>
/// Office authoring profile used to validate profile-specific markup nodes.
/// </summary>
public enum OfficeMarkupProfile {
    /// <summary>Common Markdown-compatible content without Office-specific extensions.</summary>
    Common,
    /// <summary>PowerPoint presentation authoring profile.</summary>
    Presentation,
    /// <summary>Word document authoring profile.</summary>
    Document,
    /// <summary>Excel workbook authoring profile.</summary>
    Workbook
}

/// <summary>
/// Semantic node categories exposed by the unified OfficeIMO markup AST.
/// </summary>
public enum OfficeMarkupNodeKind {
    Heading,
    Paragraph,
    List,
    Code,
    Image,
    Table,
    Diagram,
    Slide,
    PageBreak,
    Section,
    HeaderFooter,
    TableOfContents,
    Sheet,
    Range,
    Formula,
    NamedTable,
    Chart,
    Formatting,
    TextBox,
    Columns,
    Column,
    Card,
    Extension,
    RawMarkdown
}

/// <summary>
/// Severity of parser or validation diagnostics.
/// </summary>
public enum OfficeMarkupDiagnosticSeverity {
    Info,
    Warning,
    Error
}

/// <summary>
/// Parser or validation diagnostic.
/// </summary>
public sealed class OfficeMarkupDiagnostic {
    public OfficeMarkupDiagnostic(OfficeMarkupDiagnosticSeverity severity, string message, OfficeMarkupNode? node = null) {
        Severity = severity;
        Message = message ?? string.Empty;
        Node = node;
    }

    public OfficeMarkupDiagnosticSeverity Severity { get; }
    public string Message { get; }
    public OfficeMarkupNode? Node { get; }
}

/// <summary>
/// Base semantic AST node. Nodes describe office intent, not a C# or PowerShell API call.
/// </summary>
public abstract class OfficeMarkupNode {
    private readonly Dictionary<string, string> _attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    protected OfficeMarkupNode(OfficeMarkupNodeKind kind) {
        Kind = kind;
    }

    public OfficeMarkupNodeKind Kind { get; }
    public IDictionary<string, string> Attributes => _attributes;
    public string? SourceText { get; set; }
}

/// <summary>
/// Base class for block-level semantic nodes.
/// </summary>
public abstract class OfficeMarkupBlock : OfficeMarkupNode {
    protected OfficeMarkupBlock(OfficeMarkupNodeKind kind) : base(kind) {
    }
}

public sealed class OfficeMarkupPlacement {
    public string? X { get; set; }
    public string? Y { get; set; }
    public string? Width { get; set; }
    public string? Height { get; set; }

    public bool HasValue =>
        !string.IsNullOrWhiteSpace(X)
        || !string.IsNullOrWhiteSpace(Y)
        || !string.IsNullOrWhiteSpace(Width)
        || !string.IsNullOrWhiteSpace(Height);
}

/// <summary>
/// Root semantic document produced by the markup parser.
/// </summary>
public sealed class OfficeMarkupDocument {
    private readonly Dictionary<string, string> _metadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly List<OfficeMarkupBlock> _blocks = new List<OfficeMarkupBlock>();

    public OfficeMarkupDocument(OfficeMarkupProfile profile) {
        Profile = profile;
    }

    public OfficeMarkupProfile Profile { get; set; }
    public IDictionary<string, string> Metadata => _metadata;
    public IList<OfficeMarkupBlock> Blocks => _blocks;

    public IEnumerable<OfficeMarkupBlock> DescendantsAndSelf() {
        for (int i = 0; i < _blocks.Count; i++) {
            foreach (var block in EnumerateBlock(_blocks[i])) {
                yield return block;
            }
        }
    }

    private static IEnumerable<OfficeMarkupBlock> EnumerateBlock(OfficeMarkupBlock block) {
        yield return block;

        IEnumerable<OfficeMarkupBlock>? children = null;
        if (block is OfficeMarkupSlideBlock slide) {
            children = slide.Blocks;
        } else if (block is OfficeMarkupSectionBlock section) {
            children = section.Blocks;
        }

        if (children == null) {
            yield break;
        }

        foreach (var childBlock in children) {
            foreach (var child in EnumerateBlock(childBlock)) {
                yield return child;
            }
        }
    }
}

public sealed class OfficeMarkupParseResult {
    public OfficeMarkupParseResult(OfficeMarkupDocument document, IReadOnlyList<OfficeMarkupDiagnostic> diagnostics) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        Diagnostics = diagnostics ?? Array.Empty<OfficeMarkupDiagnostic>();
    }

    public OfficeMarkupDocument Document { get; }
    public IReadOnlyList<OfficeMarkupDiagnostic> Diagnostics { get; }
    public bool HasErrors => Diagnostics.Any(d => d.Severity == OfficeMarkupDiagnosticSeverity.Error);
}

public sealed class OfficeMarkupHeadingBlock : OfficeMarkupBlock {
    public OfficeMarkupHeadingBlock(int level, string text) : base(OfficeMarkupNodeKind.Heading) {
        Level = level;
        Text = text ?? string.Empty;
    }

    public int Level { get; }
    public string Text { get; }
}

public sealed class OfficeMarkupParagraphBlock : OfficeMarkupBlock {
    public OfficeMarkupParagraphBlock(string text) : base(OfficeMarkupNodeKind.Paragraph) {
        Text = text ?? string.Empty;
    }

    public string Text { get; }
}

public sealed class OfficeMarkupListBlock : OfficeMarkupBlock {
    private readonly List<OfficeMarkupListItem> _items = new List<OfficeMarkupListItem>();

    public OfficeMarkupListBlock(bool ordered, int start = 1) : base(OfficeMarkupNodeKind.List) {
        Ordered = ordered;
        Start = start;
    }

    public bool Ordered { get; }
    public int Start { get; }
    public IList<OfficeMarkupListItem> Items => _items;
}

public sealed class OfficeMarkupListItem {
    private readonly List<OfficeMarkupBlock> _blocks = new List<OfficeMarkupBlock>();

    public OfficeMarkupListItem(string text, bool isTask = false, bool isChecked = false) {
        Text = text ?? string.Empty;
        IsTask = isTask;
        IsChecked = isChecked;
    }

    public string Text { get; }
    public bool IsTask { get; }
    public bool IsChecked { get; }
    public IList<OfficeMarkupBlock> Blocks => _blocks;
}

public sealed class OfficeMarkupCodeBlock : OfficeMarkupBlock {
    public OfficeMarkupCodeBlock(string language, string content) : base(OfficeMarkupNodeKind.Code) {
        Language = language ?? string.Empty;
        Content = content ?? string.Empty;
    }

    public string Language { get; }
    public string Content { get; }
}

public sealed class OfficeMarkupImageBlock : OfficeMarkupBlock {
    public OfficeMarkupImageBlock(string source, string? alt = null, string? title = null, double? width = null, double? height = null)
        : base(OfficeMarkupNodeKind.Image) {
        Source = source ?? string.Empty;
        Alt = alt;
        Title = title;
        Width = width;
        Height = height;
    }

    public string Source { get; }
    public string? Alt { get; }
    public string? Title { get; }
    public double? Width { get; }
    public double? Height { get; }
    public OfficeMarkupPlacement? Placement { get; set; }
}

public sealed class OfficeMarkupTableBlock : OfficeMarkupBlock {
    private readonly List<string> _headers = new List<string>();
    private readonly List<IReadOnlyList<string>> _rows = new List<IReadOnlyList<string>>();

    public OfficeMarkupTableBlock() : base(OfficeMarkupNodeKind.Table) {
    }

    public IList<string> Headers => _headers;
    public IList<IReadOnlyList<string>> Rows => _rows;
}

public sealed class OfficeMarkupDiagramBlock : OfficeMarkupBlock {
    public OfficeMarkupDiagramBlock(string language, string content) : base(OfficeMarkupNodeKind.Diagram) {
        Language = language ?? string.Empty;
        Content = content ?? string.Empty;
    }

    public string Language { get; }
    public string Content { get; }
    public bool RenderAsImage { get; set; } = true;
    public OfficeMarkupPlacement? Placement { get; set; }
}

public sealed class OfficeMarkupSlideBlock : OfficeMarkupBlock {
    private readonly List<OfficeMarkupBlock> _blocks = new List<OfficeMarkupBlock>();

    public OfficeMarkupSlideBlock(string? title = null) : base(OfficeMarkupNodeKind.Slide) {
        Title = title;
    }

    public string? Title { get; set; }
    public string? Layout { get; set; }
    public string? Section { get; set; }
    public string? Transition { get; set; }
    public string? Background { get; set; }
    public string? Notes { get; set; }
    public string? Placement { get; set; }
    public int? Columns { get; set; }
    public IList<OfficeMarkupBlock> Blocks => _blocks;
}

public sealed class OfficeMarkupPageBreakBlock : OfficeMarkupBlock {
    public OfficeMarkupPageBreakBlock() : base(OfficeMarkupNodeKind.PageBreak) {
    }
}

public sealed class OfficeMarkupSectionBlock : OfficeMarkupBlock {
    private readonly List<OfficeMarkupBlock> _blocks = new List<OfficeMarkupBlock>();

    public OfficeMarkupSectionBlock(string? name = null) : base(OfficeMarkupNodeKind.Section) {
        Name = name;
    }

    public string? Name { get; set; }
    public string? PageSize { get; set; }
    public string? Orientation { get; set; }
    public IList<OfficeMarkupBlock> Blocks => _blocks;
}

public sealed class OfficeMarkupHeaderFooterBlock : OfficeMarkupBlock {
    public OfficeMarkupHeaderFooterBlock(string kind, string text) : base(OfficeMarkupNodeKind.HeaderFooter) {
        HeaderFooterKind = string.IsNullOrWhiteSpace(kind) ? "header" : kind.Trim();
        Text = text ?? string.Empty;
    }

    public string HeaderFooterKind { get; }
    public string Text { get; }
}

public sealed class OfficeMarkupTableOfContentsBlock : OfficeMarkupBlock {
    public OfficeMarkupTableOfContentsBlock() : base(OfficeMarkupNodeKind.TableOfContents) {
    }

    public int? MinLevel { get; set; }
    public int? MaxLevel { get; set; }
    public string? Title { get; set; }
}

public sealed class OfficeMarkupSheetBlock : OfficeMarkupBlock {
    public OfficeMarkupSheetBlock(string name) : base(OfficeMarkupNodeKind.Sheet) {
        Name = string.IsNullOrWhiteSpace(name) ? "Sheet1" : name.Trim();
    }

    public string Name { get; }
}

public sealed class OfficeMarkupRangeBlock : OfficeMarkupBlock {
    private readonly List<IReadOnlyList<string>> _values = new List<IReadOnlyList<string>>();

    public OfficeMarkupRangeBlock(string address) : base(OfficeMarkupNodeKind.Range) {
        Address = address ?? string.Empty;
    }

    public string Address { get; }
    public string? Sheet { get; set; }
    public IList<IReadOnlyList<string>> Values => _values;
}

public sealed class OfficeMarkupFormulaBlock : OfficeMarkupBlock {
    public OfficeMarkupFormulaBlock(string cell, string expression) : base(OfficeMarkupNodeKind.Formula) {
        Cell = cell ?? string.Empty;
        Expression = expression ?? string.Empty;
    }

    public string Cell { get; }
    public string Expression { get; }
    public string? Sheet { get; set; }
}

public sealed class OfficeMarkupNamedTableBlock : OfficeMarkupBlock {
    public OfficeMarkupNamedTableBlock(string name, string range) : base(OfficeMarkupNodeKind.NamedTable) {
        Name = name ?? string.Empty;
        Range = range ?? string.Empty;
    }

    public string Name { get; }
    public string Range { get; }
    public bool HasHeader { get; set; } = true;
}

public sealed class OfficeMarkupChartBlock : OfficeMarkupBlock {
    private readonly List<IReadOnlyList<string>> _data = new List<IReadOnlyList<string>>();

    public OfficeMarkupChartBlock(string chartType) : base(OfficeMarkupNodeKind.Chart) {
        ChartType = string.IsNullOrWhiteSpace(chartType) ? "column" : chartType.Trim();
    }

    public string ChartType { get; }
    public string? Title { get; set; }
    public string? Source { get; set; }
    public string? Sheet { get; set; }
    public OfficeMarkupPlacement? Placement { get; set; }
    public IList<IReadOnlyList<string>> Data => _data;
}

public sealed class OfficeMarkupTextBoxBlock : OfficeMarkupBlock {
    public OfficeMarkupTextBoxBlock(string text) : base(OfficeMarkupNodeKind.TextBox) {
        Text = text ?? string.Empty;
    }

    public string Text { get; }
    public string? Style { get; set; }
    public OfficeMarkupPlacement? Placement { get; set; }
}

public sealed class OfficeMarkupColumnsBlock : OfficeMarkupBlock {
    public OfficeMarkupColumnsBlock() : base(OfficeMarkupNodeKind.Columns) {
    }

    public string? Gap { get; set; }
    public OfficeMarkupPlacement? Placement { get; set; }
}

public sealed class OfficeMarkupColumnBlock : OfficeMarkupBlock {
    public OfficeMarkupColumnBlock(string columnKind, string body) : base(OfficeMarkupNodeKind.Column) {
        ColumnKind = string.IsNullOrWhiteSpace(columnKind) ? "column" : columnKind.Trim();
        Body = body ?? string.Empty;
    }

    public string ColumnKind { get; }
    public string Body { get; }
    public string? Width { get; set; }
}

public sealed class OfficeMarkupCardBlock : OfficeMarkupBlock {
    public OfficeMarkupCardBlock(string body) : base(OfficeMarkupNodeKind.Card) {
        Body = body ?? string.Empty;
    }

    public string Body { get; }
    public string? Title { get; set; }
    public string? Style { get; set; }
    public OfficeMarkupPlacement? Placement { get; set; }
}

public sealed class OfficeMarkupFormattingBlock : OfficeMarkupBlock {
    public OfficeMarkupFormattingBlock(string target) : base(OfficeMarkupNodeKind.Formatting) {
        Target = target ?? string.Empty;
    }

    public string Target { get; }
    public string? Style { get; set; }
    public string? NumberFormat { get; set; }
}

public sealed class OfficeMarkupExtensionBlock : OfficeMarkupBlock {
    private readonly ReadOnlyDictionary<string, string> _readOnlyAttributes;

    public OfficeMarkupExtensionBlock(string command, IDictionary<string, string> attributes, string body)
        : base(OfficeMarkupNodeKind.Extension) {
        Command = command ?? string.Empty;
        Body = body ?? string.Empty;
        if (attributes != null) {
            foreach (var pair in attributes) {
                Attributes[pair.Key] = pair.Value;
            }
        }

        _readOnlyAttributes = new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(Attributes, StringComparer.OrdinalIgnoreCase));
    }

    public string Command { get; }
    public string Body { get; }
    public IReadOnlyDictionary<string, string> ExtensionAttributes => _readOnlyAttributes;
}

public sealed class OfficeMarkupRawMarkdownBlock : OfficeMarkupBlock {
    public OfficeMarkupRawMarkdownBlock(string markdown) : base(OfficeMarkupNodeKind.RawMarkdown) {
        Markdown = markdown ?? string.Empty;
    }

    public string Markdown { get; }
}
