using System.Collections.ObjectModel;
using OfficeIMO.Markdown;

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
    /// <summary>A heading with a level and text.</summary>
    Heading,
    /// <summary>A text paragraph.</summary>
    Paragraph,
    /// <summary>An ordered or unordered list.</summary>
    List,
    /// <summary>A fenced code block.</summary>
    Code,
    /// <summary>An image reference.</summary>
    Image,
    /// <summary>A table of text cells.</summary>
    Table,
    /// <summary>A diagram source block.</summary>
    Diagram,
    /// <summary>A presentation slide containing blocks.</summary>
    Slide,
    /// <summary>A document page break.</summary>
    PageBreak,
    /// <summary>A document section containing blocks.</summary>
    Section,
    /// <summary>A header or footer text block.</summary>
    HeaderFooter,
    /// <summary>A table-of-contents request.</summary>
    TableOfContents,
    /// <summary>A workbook sheet declaration.</summary>
    Sheet,
    /// <summary>A workbook cell range and its values.</summary>
    Range,
    /// <summary>A cell formula.</summary>
    Formula,
    /// <summary>A named workbook table.</summary>
    NamedTable,
    /// <summary>A chart definition.</summary>
    Chart,
    /// <summary>Formatting applied to a named target.</summary>
    Formatting,
    /// <summary>A positioned text box.</summary>
    TextBox,
    /// <summary>A column layout request.</summary>
    Columns,
    /// <summary>One column in a column layout.</summary>
    Column,
    /// <summary>A card with optional title and style.</summary>
    Card,
    /// <summary>An Office-aware directive not mapped to a built-in node type.</summary>
    Extension,
    /// <summary>Markdown retained without a semantic mapping.</summary>
    RawMarkdown
}

/// <summary>
/// Severity of parser or validation diagnostics.
/// </summary>
public enum OfficeMarkupDiagnosticSeverity {
    /// <summary>Informational parser or validation feedback.</summary>
    Info,
    /// <summary>A potential authoring problem that does not count as an error.</summary>
    Warning,
    /// <summary>An authoring problem counted by <see cref="OfficeMarkupParseResult.HasErrors"/>.</summary>
    Error
}

/// <summary>
/// Parser or validation diagnostic.
/// </summary>
public sealed class OfficeMarkupDiagnostic {
    /// <summary>Creates a diagnostic, optionally associated with a semantic node.</summary>
    public OfficeMarkupDiagnostic(OfficeMarkupDiagnosticSeverity severity, string message, OfficeMarkupNode? node = null)
        : this(severity, message, node, InferLossKind(severity)) {
    }

    /// <summary>Creates a diagnostic with an explicit fidelity-loss category.</summary>
    public OfficeMarkupDiagnostic(
        OfficeMarkupDiagnosticSeverity severity,
        string message,
        OfficeMarkupNode? node,
        OfficeConversionLossKind lossKind) {
        Severity = severity;
        Message = message ?? string.Empty;
        LossKind = lossKind;
        Node = node;
    }

    /// <summary>Gets the impact assigned by the parser or validator.</summary>
    public OfficeMarkupDiagnosticSeverity Severity { get; }
    /// <summary>Gets the diagnostic text; a null constructor value becomes empty.</summary>
    public string Message { get; }
    /// <summary>Gets the exact fidelity-loss category represented by the diagnostic.</summary>
    public OfficeConversionLossKind LossKind { get; }
    /// <summary>Gets the related node when one was supplied.</summary>
    public OfficeMarkupNode? Node { get; }

    private static OfficeConversionLossKind InferLossKind(OfficeMarkupDiagnosticSeverity severity) => severity switch {
        OfficeMarkupDiagnosticSeverity.Info => OfficeConversionLossKind.None,
        OfficeMarkupDiagnosticSeverity.Warning => OfficeConversionLossKind.Approximation,
        _ => OfficeConversionLossKind.Failure
    };
}

/// <summary>
/// Base semantic AST node. Nodes describe office intent, not a C# or PowerShell API call.
/// </summary>
public abstract class OfficeMarkupNode {
    private readonly Dictionary<string, string> _attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private string? _sourceText;
    private IMarkdownBlock? _sourceMarkdownBlock;

    /// <summary>Initializes the node's semantic category.</summary>
    protected OfficeMarkupNode(OfficeMarkupNodeKind kind) {
        Kind = kind;
    }

    /// <summary>Gets the semantic category of this node.</summary>
    public OfficeMarkupNodeKind Kind { get; }
    /// <summary>Gets mutable, case-insensitively keyed directive attributes.</summary>
    public IDictionary<string, string> Attributes => _attributes;
    /// <summary>Gets or sets the original markup for this node, when retained.</summary>
    /// <remarks>For parser-produced nodes, the getter may render a retained Markdown block on first access.</remarks>
    public string? SourceText {
        get {
            if (_sourceText == null && _sourceMarkdownBlock != null) {
                _sourceText = _sourceMarkdownBlock.RenderMarkdown();
                _sourceMarkdownBlock = null;
            }
            return _sourceText;
        }
        set {
            _sourceText = value;
            _sourceMarkdownBlock = null;
        }
    }

    internal void SetLazySourceText(IMarkdownBlock sourceMarkdownBlock) {
        _sourceText = null;
        _sourceMarkdownBlock = sourceMarkdownBlock ?? throw new ArgumentNullException(nameof(sourceMarkdownBlock));
    }
}

/// <summary>
/// Base class for block-level semantic nodes.
/// </summary>
public abstract class OfficeMarkupBlock : OfficeMarkupNode {
    /// <summary>Initializes a block with its semantic category.</summary>
    protected OfficeMarkupBlock(OfficeMarkupNodeKind kind) : base(kind) {
    }
}

/// <summary>Optional placement values retained as markup text for target-specific interpretation.</summary>
public sealed class OfficeMarkupPlacement {
    /// <summary>Gets or sets the horizontal-position expression.</summary>
    public string? X { get; set; }
    /// <summary>Gets or sets the vertical-position expression.</summary>
    public string? Y { get; set; }
    /// <summary>Gets or sets the width expression.</summary>
    public string? Width { get; set; }
    /// <summary>Gets or sets the height expression.</summary>
    public string? Height { get; set; }

    /// <summary>Gets whether any placement expression is nonblank.</summary>
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

    /// <summary>Creates an empty semantic document for the selected authoring profile.</summary>
    public OfficeMarkupDocument(OfficeMarkupProfile profile) {
        Profile = profile;
    }

    /// <summary>Gets or sets the authoring profile used by validation and target-specific consumers.</summary>
    public OfficeMarkupProfile Profile { get; set; }
    /// <summary>Gets mutable, case-insensitively keyed front-matter metadata.</summary>
    public IDictionary<string, string> Metadata => _metadata;
    /// <summary>Gets the mutable top-level block sequence.</summary>
    public IList<OfficeMarkupBlock> Blocks => _blocks;

    /// <summary>Enumerates top-level blocks and blocks nested in slides or sections in depth-first order.</summary>
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

/// <summary>A semantic document and the diagnostics produced while parsing or validating it.</summary>
public sealed class OfficeMarkupParseResult {
    /// <summary>Creates a parse result, treating null diagnostics as an empty sequence.</summary>
    public OfficeMarkupParseResult(OfficeMarkupDocument document, IReadOnlyList<OfficeMarkupDiagnostic> diagnostics) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        Diagnostics = diagnostics ?? Array.Empty<OfficeMarkupDiagnostic>();
    }

    /// <summary>Gets the semantic document; its metadata and blocks remain mutable.</summary>
    public OfficeMarkupDocument Document { get; }
    /// <summary>Gets the supplied diagnostics or an empty sequence when null was supplied.</summary>
    public IReadOnlyList<OfficeMarkupDiagnostic> Diagnostics { get; }
    /// <summary>Gets whether any diagnostic has error severity.</summary>
    public bool HasErrors => Diagnostics.Any(d => d.Severity == OfficeMarkupDiagnosticSeverity.Error);
}

/// <summary>A heading with a level and text.</summary>
public sealed class OfficeMarkupHeadingBlock : OfficeMarkupBlock {
    /// <summary>Creates a heading; the supplied level is retained without normalization.</summary>
    public OfficeMarkupHeadingBlock(int level, string text) : base(OfficeMarkupNodeKind.Heading) {
        Level = level;
        Text = text ?? string.Empty;
    }

    /// <summary>Gets the supplied heading level.</summary>
    public int Level { get; }
    /// <summary>Gets the heading text, or an empty string when null was supplied.</summary>
    public string Text { get; }
}

/// <summary>A paragraph of markup text.</summary>
public sealed class OfficeMarkupParagraphBlock : OfficeMarkupBlock {
    /// <summary>Creates a paragraph, treating null text as empty.</summary>
    public OfficeMarkupParagraphBlock(string text) : base(OfficeMarkupNodeKind.Paragraph) {
        Text = text ?? string.Empty;
    }

    /// <summary>Gets the paragraph text.</summary>
    public string Text { get; }
}

/// <summary>An ordered or unordered list of semantic list items.</summary>
public sealed class OfficeMarkupListBlock : OfficeMarkupBlock {
    private readonly List<OfficeMarkupListItem> _items = new List<OfficeMarkupListItem>();

    /// <summary>Creates a list with its ordering flag and starting number.</summary>
    public OfficeMarkupListBlock(bool ordered, int start = 1) : base(OfficeMarkupNodeKind.List) {
        Ordered = ordered;
        Start = start;
    }

    /// <summary>Gets whether the list is ordered.</summary>
    public bool Ordered { get; }
    /// <summary>Gets the supplied starting number; unordered lists retain it too.</summary>
    public int Start { get; }
    /// <summary>Gets the mutable list of items.</summary>
    public IList<OfficeMarkupListItem> Items => _items;
}

/// <summary>A list item with optional task state and nested blocks.</summary>
public sealed class OfficeMarkupListItem {
    private readonly List<OfficeMarkupBlock> _blocks = new List<OfficeMarkupBlock>();

    /// <summary>Creates a list item, treating null text as empty.</summary>
    public OfficeMarkupListItem(string text, bool isTask = false, bool isChecked = false) {
        Text = text ?? string.Empty;
        IsTask = isTask;
        IsChecked = isChecked;
    }

    /// <summary>Gets the item's text.</summary>
    public string Text { get; }
    /// <summary>Gets whether the item represents a task.</summary>
    public bool IsTask { get; }
    /// <summary>Gets the supplied checked state.</summary>
    public bool IsChecked { get; }
    /// <summary>Gets mutable blocks nested under the item.</summary>
    public IList<OfficeMarkupBlock> Blocks => _blocks;
}

/// <summary>A fenced code block with a language label and source content.</summary>
public sealed class OfficeMarkupCodeBlock : OfficeMarkupBlock {
    /// <summary>Creates a code block, treating null language or content as empty.</summary>
    public OfficeMarkupCodeBlock(string language, string content) : base(OfficeMarkupNodeKind.Code) {
        Language = language ?? string.Empty;
        Content = content ?? string.Empty;
    }

    /// <summary>Gets the supplied language label.</summary>
    public string Language { get; }
    /// <summary>Gets the code content.</summary>
    public string Content { get; }
}

/// <summary>An image reference with optional descriptive text, size values, and placement.</summary>
public sealed class OfficeMarkupImageBlock : OfficeMarkupBlock {
    /// <summary>Creates an image reference without loading the referenced content.</summary>
    public OfficeMarkupImageBlock(string source, string? alt = null, string? title = null, double? width = null, double? height = null)
        : base(OfficeMarkupNodeKind.Image) {
        Source = source ?? string.Empty;
        Alt = alt;
        Title = title;
        Width = width;
        Height = height;
    }

    /// <summary>Gets the image source string, or empty when null was supplied.</summary>
    public string Source { get; }
    /// <summary>Gets optional alternative text.</summary>
    public string? Alt { get; }
    /// <summary>Gets an optional image title.</summary>
    public string? Title { get; }
    /// <summary>Gets an optional numeric width supplied by the parser or caller.</summary>
    public double? Width { get; }
    /// <summary>Gets an optional numeric height supplied by the parser or caller.</summary>
    public double? Height { get; }
    /// <summary>Gets or sets optional placement expressions for a target renderer.</summary>
    public OfficeMarkupPlacement? Placement { get; set; }
}

/// <summary>A table represented by header strings and rows of string cells.</summary>
public sealed class OfficeMarkupTableBlock : OfficeMarkupBlock {
    private readonly List<string> _headers = new List<string>();
    private readonly List<IReadOnlyList<string>> _rows = new List<IReadOnlyList<string>>();

    /// <summary>Creates an empty table.</summary>
    public OfficeMarkupTableBlock() : base(OfficeMarkupNodeKind.Table) {
    }

    /// <summary>Gets the mutable header cells.</summary>
    public IList<string> Headers => _headers;
    /// <summary>Gets the mutable row sequence; each row is exposed as a read-only list.</summary>
    public IList<IReadOnlyList<string>> Rows => _rows;
}

/// <summary>A diagram source block with a renderer hint.</summary>
public sealed class OfficeMarkupDiagramBlock : OfficeMarkupBlock {
    /// <summary>Creates a diagram block, treating null language or content as empty.</summary>
    public OfficeMarkupDiagramBlock(string language, string content) : base(OfficeMarkupNodeKind.Diagram) {
        Language = language ?? string.Empty;
        Content = content ?? string.Empty;
    }

    /// <summary>Gets the diagram language label.</summary>
    public string Language { get; }
    /// <summary>Gets the diagram source text.</summary>
    public string Content { get; }
    /// <summary>Gets or sets whether consumers should render the diagram as an image; defaults to true.</summary>
    public bool RenderAsImage { get; set; } = true;
    /// <summary>Gets or sets optional placement expressions for a target renderer.</summary>
    public OfficeMarkupPlacement? Placement { get; set; }
}

/// <summary>A presentation slide and its nested semantic blocks.</summary>
public sealed class OfficeMarkupSlideBlock : OfficeMarkupBlock {
    private readonly List<OfficeMarkupBlock> _blocks = new List<OfficeMarkupBlock>();

    /// <summary>Creates a slide with an optional title.</summary>
    public OfficeMarkupSlideBlock(string? title = null) : base(OfficeMarkupNodeKind.Slide) {
        Title = title;
    }

    /// <summary>Gets or sets the slide title.</summary>
    public string? Title { get; set; }
    /// <summary>Gets or sets the requested slide layout name.</summary>
    public string? Layout { get; set; }
    /// <summary>Gets or sets the associated presentation section name.</summary>
    public string? Section { get; set; }
    /// <summary>Gets or sets the transition expression for a target renderer.</summary>
    public string? Transition { get; set; }
    /// <summary>Gets or sets the slide background expression.</summary>
    public string? Background { get; set; }
    /// <summary>Gets or sets speaker-notes text.</summary>
    public string? Notes { get; set; }
    /// <summary>Gets or sets the slide's placement hint.</summary>
    public string? Placement { get; set; }
    /// <summary>Gets or sets an optional column-count hint.</summary>
    public int? Columns { get; set; }
    /// <summary>Gets the mutable sequence of slide blocks.</summary>
    public IList<OfficeMarkupBlock> Blocks => _blocks;
}

/// <summary>A document page-break request.</summary>
public sealed class OfficeMarkupPageBreakBlock : OfficeMarkupBlock {
    /// <summary>Creates a page-break block.</summary>
    public OfficeMarkupPageBreakBlock() : base(OfficeMarkupNodeKind.PageBreak) {
    }
}

/// <summary>A document section and its nested semantic blocks.</summary>
public sealed class OfficeMarkupSectionBlock : OfficeMarkupBlock {
    private readonly List<OfficeMarkupBlock> _blocks = new List<OfficeMarkupBlock>();

    /// <summary>Creates a section with an optional name.</summary>
    public OfficeMarkupSectionBlock(string? name = null) : base(OfficeMarkupNodeKind.Section) {
        Name = name;
    }

    /// <summary>Gets or sets the section name.</summary>
    public string? Name { get; set; }
    /// <summary>Gets or sets the requested page-size expression.</summary>
    public string? PageSize { get; set; }
    /// <summary>Gets or sets the requested page orientation.</summary>
    public string? Orientation { get; set; }
    /// <summary>Gets the mutable sequence of section blocks.</summary>
    public IList<OfficeMarkupBlock> Blocks => _blocks;
}

/// <summary>Text requested for a document header or footer.</summary>
public sealed class OfficeMarkupHeaderFooterBlock : OfficeMarkupBlock {
    /// <summary>Creates a header or footer block; blank kind becomes <c>header</c>.</summary>
    public OfficeMarkupHeaderFooterBlock(string kind, string text) : base(OfficeMarkupNodeKind.HeaderFooter) {
        HeaderFooterKind = string.IsNullOrWhiteSpace(kind) ? "header" : kind.Trim();
        Text = text ?? string.Empty;
    }

    /// <summary>Gets the supplied kind after trimming, or <c>header</c> for blank input.</summary>
    public string HeaderFooterKind { get; }
    /// <summary>Gets the text, or empty when null was supplied.</summary>
    public string Text { get; }
}

/// <summary>A table-of-contents request with optional level bounds and title.</summary>
public sealed class OfficeMarkupTableOfContentsBlock : OfficeMarkupBlock {
    /// <summary>Creates a table-of-contents block with no level bounds.</summary>
    public OfficeMarkupTableOfContentsBlock() : base(OfficeMarkupNodeKind.TableOfContents) {
    }

    /// <summary>Gets or sets the optional minimum heading level.</summary>
    public int? MinLevel { get; set; }
    /// <summary>Gets or sets the optional maximum heading level.</summary>
    public int? MaxLevel { get; set; }
    /// <summary>Gets or sets an optional heading for the table of contents.</summary>
    public string? Title { get; set; }
}

/// <summary>A workbook sheet declaration.</summary>
public sealed class OfficeMarkupSheetBlock : OfficeMarkupBlock {
    /// <summary>Creates a sheet; blank names become <c>Sheet1</c>.</summary>
    public OfficeMarkupSheetBlock(string name) : base(OfficeMarkupNodeKind.Sheet) {
        Name = string.IsNullOrWhiteSpace(name) ? "Sheet1" : name.Trim();
    }

    /// <summary>Gets the trimmed name, or <c>Sheet1</c> when blank was supplied.</summary>
    public string Name { get; }
}

/// <summary>A workbook range address and rows of cell values.</summary>
public sealed class OfficeMarkupRangeBlock : OfficeMarkupBlock {
    private readonly List<IReadOnlyList<string>> _values = new List<IReadOnlyList<string>>();

    /// <summary>Creates a range block, treating a null address as empty.</summary>
    public OfficeMarkupRangeBlock(string address) : base(OfficeMarkupNodeKind.Range) {
        Address = address ?? string.Empty;
    }

    /// <summary>Gets the supplied cell-range address.</summary>
    public string Address { get; }
    /// <summary>Gets or sets an optional sheet name.</summary>
    public string? Sheet { get; set; }
    /// <summary>Gets the mutable row sequence of cell values.</summary>
    public IList<IReadOnlyList<string>> Values => _values;
}

/// <summary>A formula expression associated with a workbook cell.</summary>
public sealed class OfficeMarkupFormulaBlock : OfficeMarkupBlock {
    /// <summary>Creates a formula block, treating null cell or expression values as empty.</summary>
    public OfficeMarkupFormulaBlock(string cell, string expression) : base(OfficeMarkupNodeKind.Formula) {
        Cell = cell ?? string.Empty;
        Expression = expression ?? string.Empty;
    }

    /// <summary>Gets the target cell address.</summary>
    public string Cell { get; }
    /// <summary>Gets the formula expression as supplied.</summary>
    public string Expression { get; }
    /// <summary>Gets or sets an optional sheet name.</summary>
    public string? Sheet { get; set; }
}

/// <summary>A named workbook table over a range.</summary>
public sealed class OfficeMarkupNamedTableBlock : OfficeMarkupBlock {
    /// <summary>Creates a named-table block, treating null name or range values as empty.</summary>
    public OfficeMarkupNamedTableBlock(string name, string range) : base(OfficeMarkupNodeKind.NamedTable) {
        Name = name ?? string.Empty;
        Range = range ?? string.Empty;
    }

    /// <summary>Gets the table name.</summary>
    public string Name { get; }
    /// <summary>Gets the table range expression.</summary>
    public string Range { get; }
    /// <summary>Gets or sets whether the first row is a header; defaults to true.</summary>
    public bool HasHeader { get; set; } = true;
}

/// <summary>A chart request with optional source range or inline data.</summary>
public sealed class OfficeMarkupChartBlock : OfficeMarkupBlock {
    private readonly List<IReadOnlyList<string>> _data = new List<IReadOnlyList<string>>();

    /// <summary>Creates a chart; blank chart types become <c>column</c>.</summary>
    public OfficeMarkupChartBlock(string chartType) : base(OfficeMarkupNodeKind.Chart) {
        ChartType = string.IsNullOrWhiteSpace(chartType) ? "column" : chartType.Trim();
    }

    /// <summary>Gets the trimmed chart-type name, or <c>column</c> for blank input.</summary>
    public string ChartType { get; }
    /// <summary>Gets or sets an optional chart title.</summary>
    public string? Title { get; set; }
    /// <summary>Gets or sets an optional chart data-source expression.</summary>
    public string? Source { get; set; }
    /// <summary>Gets or sets an optional sheet name.</summary>
    public string? Sheet { get; set; }
    /// <summary>Gets or sets optional placement expressions for a target renderer.</summary>
    public OfficeMarkupPlacement? Placement { get; set; }
    /// <summary>Gets the mutable sequence of inline chart-data rows.</summary>
    public IList<IReadOnlyList<string>> Data => _data;
}

/// <summary>Text requested in an optionally positioned text box.</summary>
public sealed class OfficeMarkupTextBoxBlock : OfficeMarkupBlock {
    /// <summary>Creates a text box, treating null text as empty.</summary>
    public OfficeMarkupTextBoxBlock(string text) : base(OfficeMarkupNodeKind.TextBox) {
        Text = text ?? string.Empty;
    }

    /// <summary>Gets the text-box content.</summary>
    public string Text { get; }
    /// <summary>Gets or sets an optional style name or expression.</summary>
    public string? Style { get; set; }
    /// <summary>Gets or sets optional placement expressions for a target renderer.</summary>
    public OfficeMarkupPlacement? Placement { get; set; }
}

/// <summary>A column-layout request with optional spacing and placement.</summary>
public sealed class OfficeMarkupColumnsBlock : OfficeMarkupBlock {
    /// <summary>Creates a column-layout block.</summary>
    public OfficeMarkupColumnsBlock() : base(OfficeMarkupNodeKind.Columns) {
    }

    /// <summary>Gets or sets the gap expression between columns.</summary>
    public string? Gap { get; set; }
    /// <summary>Gets or sets optional placement expressions for a target renderer.</summary>
    public OfficeMarkupPlacement? Placement { get; set; }
}

/// <summary>One column's kind, body text, and optional width.</summary>
public sealed class OfficeMarkupColumnBlock : OfficeMarkupBlock {
    /// <summary>Creates a column; blank kinds become <c>column</c>.</summary>
    public OfficeMarkupColumnBlock(string columnKind, string body) : base(OfficeMarkupNodeKind.Column) {
        ColumnKind = string.IsNullOrWhiteSpace(columnKind) ? "column" : columnKind.Trim();
        Body = body ?? string.Empty;
    }

    /// <summary>Gets the trimmed kind, or <c>column</c> for blank input.</summary>
    public string ColumnKind { get; }
    /// <summary>Gets the column body, or empty when null was supplied.</summary>
    public string Body { get; }
    /// <summary>Gets or sets an optional width expression.</summary>
    public string? Width { get; set; }
}

/// <summary>A card with body text and optional presentation hints.</summary>
public sealed class OfficeMarkupCardBlock : OfficeMarkupBlock {
    /// <summary>Creates a card, treating null body text as empty.</summary>
    public OfficeMarkupCardBlock(string body) : base(OfficeMarkupNodeKind.Card) {
        Body = body ?? string.Empty;
    }

    /// <summary>Gets the card body.</summary>
    public string Body { get; }
    /// <summary>Gets or sets an optional card title.</summary>
    public string? Title { get; set; }
    /// <summary>Gets or sets an optional style name or expression.</summary>
    public string? Style { get; set; }
    /// <summary>Gets or sets optional placement expressions for a target renderer.</summary>
    public OfficeMarkupPlacement? Placement { get; set; }
}

/// <summary>Formatting instructions for a named semantic target.</summary>
public sealed class OfficeMarkupFormattingBlock : OfficeMarkupBlock {
    /// <summary>Creates a formatting block, treating a null target as empty.</summary>
    public OfficeMarkupFormattingBlock(string target) : base(OfficeMarkupNodeKind.Formatting) {
        Target = target ?? string.Empty;
    }

    /// <summary>Gets the supplied target expression.</summary>
    public string Target { get; }
    /// <summary>Gets or sets an optional style name or expression.</summary>
    public string? Style { get; set; }
    /// <summary>Gets or sets an optional number-format expression.</summary>
    public string? NumberFormat { get; set; }
}

/// <summary>An unrecognized or extensible Office directive retained for downstream consumers.</summary>
public sealed class OfficeMarkupExtensionBlock : OfficeMarkupBlock {
    private readonly ReadOnlyDictionary<string, string> _readOnlyAttributes;

    /// <summary>Creates an extension block and snapshots the supplied attributes.</summary>
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

    /// <summary>Gets the directive command, or empty when null was supplied.</summary>
    public string Command { get; }
    /// <summary>Gets the directive body, or empty when null was supplied.</summary>
    public string Body { get; }
    /// <summary>Gets the read-only snapshot of attributes supplied at construction.</summary>
    /// <remarks>The inherited <see cref="OfficeMarkupNode.Attributes"/> dictionary remains independently mutable.</remarks>
    public IReadOnlyDictionary<string, string> ExtensionAttributes => _readOnlyAttributes;
}

/// <summary>Markdown retained as text for a downstream renderer.</summary>
public sealed class OfficeMarkupRawMarkdownBlock : OfficeMarkupBlock {
    /// <summary>Creates a raw-Markdown block, treating null text as empty.</summary>
    public OfficeMarkupRawMarkdownBlock(string markdown) : base(OfficeMarkupNodeKind.RawMarkdown) {
        Markdown = markdown ?? string.Empty;
    }

    /// <summary>Gets the retained Markdown text.</summary>
    public string Markdown { get; }
}
