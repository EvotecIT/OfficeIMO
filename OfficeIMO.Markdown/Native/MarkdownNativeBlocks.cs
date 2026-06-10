namespace OfficeIMO.Markdown;

/// <summary>
/// Base type for an AST-backed native markdown block projection.
/// </summary>
public abstract class MarkdownNativeBlock {
    private protected MarkdownNativeBlock(
        MarkdownNativeBlockKind kind,
        IMarkdownBlock sourceBlock,
        MarkdownSyntaxNode syntaxNode) {
        Kind = kind;
        SourceBlock = sourceBlock ?? throw new ArgumentNullException(nameof(sourceBlock));
        SyntaxNode = syntaxNode ?? throw new ArgumentNullException(nameof(syntaxNode));
        SourceSpan = syntaxNode.SourceSpan ?? (sourceBlock as MarkdownObject)?.SourceSpan;
        Id = MarkdownNativeBlockId.Create(kind, sourceBlock, syntaxNode, SourceSpan);
    }

    /// <summary>Deterministic identity for this projection within stable markdown input.</summary>
    public string Id { get; }

    /// <summary>Native projection kind.</summary>
    public MarkdownNativeBlockKind Kind { get; }

    /// <summary>Source span in the normalized markdown text when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Syntax node that produced this native block.</summary>
    public MarkdownSyntaxNode SyntaxNode { get; }

    /// <summary>Original OfficeIMO markdown block backing this projection.</summary>
    public IMarkdownBlock SourceBlock { get; }

    /// <summary>Returns <see langword="true"/> when this block's source span contains the supplied 1-based line.</summary>
    public bool ContainsLine(int lineNumber) => SourceSpan.HasValue && SourceSpan.Value.ContainsLine(lineNumber);
}

/// <summary>
/// Native projection for a paragraph block.
/// </summary>
public sealed class MarkdownNativeParagraphBlock : MarkdownNativeBlock {
    internal MarkdownNativeParagraphBlock(ParagraphBlock paragraph, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Paragraph, paragraph, syntaxNode) {
        Paragraph = paragraph;
        Inlines = paragraph.Inlines;
        Text = InlinePlainText.Extract(paragraph.Inlines);
    }

    /// <summary>Source paragraph block.</summary>
    public ParagraphBlock Paragraph { get; }

    /// <summary>Plain-text paragraph content.</summary>
    public string Text { get; }

    /// <summary>Structured inline nodes.</summary>
    public InlineSequence Inlines { get; }
}

/// <summary>
/// Native projection for a code block.
/// </summary>
public sealed class MarkdownNativeCodeBlock : MarkdownNativeBlock {
    internal MarkdownNativeCodeBlock(CodeBlock code, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Code, code, syntaxNode) {
        Code = code;
        Language = code.Language;
        InfoString = code.InfoString;
        FenceInfo = code.FenceInfo;
        Content = code.Content;
        Caption = code.Caption;
        Attributes = code.FenceInfo.Attributes;
        Classes = code.FenceInfo.Classes;
        ElementId = code.FenceInfo.ElementId;
        Title = code.FenceInfo.Title;
    }

    /// <summary>Source code block.</summary>
    public CodeBlock Code { get; }

    /// <summary>Primary fence language token.</summary>
    public string Language { get; }

    /// <summary>Full fenced-code info string.</summary>
    public string InfoString { get; }

    /// <summary>Structured fenced-code metadata.</summary>
    public MarkdownCodeFenceInfo FenceInfo { get; }

    /// <summary>Code content with normalized line endings.</summary>
    public string Content { get; }

    /// <summary>Optional code-block caption.</summary>
    public string? Caption { get; }

    /// <summary>Parsed fence attributes.</summary>
    public IReadOnlyDictionary<string, string?> Attributes { get; }

    /// <summary>Parsed fence classes.</summary>
    public IReadOnlyList<string> Classes { get; }

    /// <summary>Parsed fence element id.</summary>
    public string? ElementId { get; }

    /// <summary>Convenience title resolved from fence metadata.</summary>
    public string? Title { get; }
}

/// <summary>
/// Native projection for a semantic visual fenced block.
/// </summary>
public sealed class MarkdownNativeVisualBlock : MarkdownNativeBlock {
    internal MarkdownNativeVisualBlock(SemanticFencedBlock visual, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Visual, visual, syntaxNode) {
        Visual = visual;
        SemanticKind = visual.SemanticKind;
        Language = visual.Language;
        InfoString = visual.InfoString;
        FenceInfo = visual.FenceInfo;
        Content = visual.Content;
        Caption = visual.Caption;
        Attributes = visual.FenceInfo.Attributes;
        Classes = visual.FenceInfo.Classes;
        ElementId = visual.FenceInfo.ElementId;
        Title = visual.FenceInfo.Title;
    }

    /// <summary>Source semantic fenced block.</summary>
    public SemanticFencedBlock Visual { get; }

    /// <summary>Host-defined semantic kind such as chart, network, dataview, or mermaid.</summary>
    public string SemanticKind { get; }

    /// <summary>Primary fence language token.</summary>
    public string Language { get; }

    /// <summary>Full fenced-code info string.</summary>
    public string InfoString { get; }

    /// <summary>Structured fenced-code metadata.</summary>
    public MarkdownCodeFenceInfo FenceInfo { get; }

    /// <summary>Visual payload with normalized line endings.</summary>
    public string Content { get; }

    /// <summary>Optional visual-block caption.</summary>
    public string? Caption { get; }

    /// <summary>Parsed fence attributes.</summary>
    public IReadOnlyDictionary<string, string?> Attributes { get; }

    /// <summary>Parsed fence classes.</summary>
    public IReadOnlyList<string> Classes { get; }

    /// <summary>Parsed fence element id.</summary>
    public string? ElementId { get; }

    /// <summary>Convenience title resolved from fence metadata.</summary>
    public string? Title { get; }
}

/// <summary>
/// Native projection for a markdown table.
/// </summary>
public sealed class MarkdownNativeTableBlock : MarkdownNativeBlock {
    internal MarkdownNativeTableBlock(TableBlock table, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Table, table, syntaxNode) {
        Table = table;
        HeaderCells = BuildHeaderCells(table, syntaxNode);
        Rows = BuildRows(table, syntaxNode);
    }

    /// <summary>Source table block.</summary>
    public TableBlock Table { get; }

    /// <summary>Header cells in document order.</summary>
    public IReadOnlyList<MarkdownNativeTableCell> HeaderCells { get; }

    /// <summary>Body rows and cells in document order.</summary>
    public IReadOnlyList<IReadOnlyList<MarkdownNativeTableCell>> Rows { get; }

    private static IReadOnlyList<MarkdownNativeTableCell> BuildHeaderCells(TableBlock table, MarkdownSyntaxNode syntaxNode) {
        if (table.Headers.Count == 0) {
            return Array.Empty<MarkdownNativeTableCell>();
        }

        var headerNode = syntaxNode.Children.FirstOrDefault(static child => child.Kind == MarkdownSyntaxKind.TableHeader);
        return BuildCells(table.Headers, table.HeaderCells, table.Alignments, headerNode, isHeader: true, rowIndex: -1);
    }

    private static IReadOnlyList<IReadOnlyList<MarkdownNativeTableCell>> BuildRows(TableBlock table, MarkdownSyntaxNode syntaxNode) {
        if (table.Rows.Count == 0 && table.RowCells.Count == 0) {
            return Array.Empty<IReadOnlyList<MarkdownNativeTableCell>>();
        }

        var rowNodes = syntaxNode.Children.Where(static child => child.Kind == MarkdownSyntaxKind.TableRow).ToArray();
        var rows = new List<IReadOnlyList<MarkdownNativeTableCell>>(Math.Max(table.Rows.Count, table.RowCells.Count));
        var rowCount = Math.Max(table.Rows.Count, table.RowCells.Count);
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            var rawCells = rowIndex < table.Rows.Count ? table.Rows[rowIndex] : Array.Empty<string>();
            var structuredCells = rowIndex < table.RowCells.Count ? table.RowCells[rowIndex] : Array.Empty<TableCell>();
            var rowNode = rowIndex < rowNodes.Length ? rowNodes[rowIndex] : null;
            rows.Add(BuildCells(rawCells, structuredCells, table.Alignments, rowNode, isHeader: false, rowIndex: rowIndex));
        }

        return rows;
    }

    private static IReadOnlyList<MarkdownNativeTableCell> BuildCells(
        IReadOnlyList<string> rawCells,
        IReadOnlyList<TableCell> structuredCells,
        IReadOnlyList<ColumnAlignment> columnAlignments,
        MarkdownSyntaxNode? rowNode,
        bool isHeader,
        int rowIndex) {
        var count = Math.Max(rawCells?.Count ?? 0, structuredCells?.Count ?? 0);
        if (count == 0) {
            return Array.Empty<MarkdownNativeTableCell>();
        }

        var cells = new List<MarkdownNativeTableCell>(count);
        for (var columnIndex = 0; columnIndex < count; columnIndex++) {
            var raw = rawCells != null && columnIndex < rawCells.Count ? rawCells[columnIndex] ?? string.Empty : string.Empty;
            var cell = structuredCells != null && columnIndex < structuredCells.Count ? structuredCells[columnIndex] : null;
            var cellNode = rowNode != null && columnIndex < rowNode.Children.Count ? rowNode.Children[columnIndex] : null;
            var alignment = ResolveAlignment(cell, columnAlignments, columnIndex);
            cells.Add(new MarkdownNativeTableCell(raw, cell, cellNode, isHeader, rowIndex, columnIndex, alignment));
        }

        return cells;
    }

    private static ColumnAlignment ResolveAlignment(
        TableCell? sourceCell,
        IReadOnlyList<ColumnAlignment> columnAlignments,
        int columnIndex) {
        if (sourceCell != null && sourceCell.Alignment != ColumnAlignment.None) {
            return sourceCell.Alignment;
        }

        return columnAlignments != null && columnIndex >= 0 && columnIndex < columnAlignments.Count
            ? columnAlignments[columnIndex]
            : ColumnAlignment.None;
    }
}

/// <summary>
/// Native projection for a table cell.
/// </summary>
public sealed class MarkdownNativeTableCell {
    internal MarkdownNativeTableCell(
        string rawText,
        TableCell? sourceCell,
        MarkdownSyntaxNode? syntaxNode,
        bool isHeader,
        int rowIndex,
        int columnIndex,
        ColumnAlignment alignment) {
        RawText = rawText ?? string.Empty;
        SourceCell = sourceCell;
        SyntaxNode = syntaxNode;
        SourceSpan = syntaxNode?.SourceSpan ?? sourceCell?.SourceSpan;
        IsHeader = isHeader;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Text = ExtractText(sourceCell, RawText);
        Markdown = sourceCell?.Markdown ?? RawText;
        Blocks = sourceCell != null ? sourceCell.Blocks : Array.Empty<IMarkdownBlock>();
        Alignment = alignment;
    }

    /// <summary>Raw cell text from the table source.</summary>
    public string RawText { get; }

    /// <summary>Plain-text cell content.</summary>
    public string Text { get; }

    /// <summary>Markdown representation of the cell content.</summary>
    public string Markdown { get; }

    /// <summary>Structured child blocks in the cell.</summary>
    public IReadOnlyList<IMarkdownBlock> Blocks { get; }

    /// <summary>Source table cell when structured cell data is available.</summary>
    public TableCell? SourceCell { get; }

    /// <summary>Syntax node that produced this cell when available.</summary>
    public MarkdownSyntaxNode? SyntaxNode { get; }

    /// <summary>Source span in the normalized markdown text when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Whether this cell belongs to the table header.</summary>
    public bool IsHeader { get; }

    /// <summary>Zero-based data row index, or <c>-1</c> for header cells.</summary>
    public int RowIndex { get; }

    /// <summary>Zero-based column index.</summary>
    public int ColumnIndex { get; }

    /// <summary>Cell-level alignment override when present.</summary>
    public ColumnAlignment Alignment { get; }

    private static string ExtractText(TableCell? sourceCell, string rawText) {
        if (sourceCell == null || sourceCell.Blocks.Count == 0) {
            return rawText ?? string.Empty;
        }

        if (sourceCell.Blocks.Count == 1 && sourceCell.Blocks[0] is ParagraphBlock paragraph) {
            return InlinePlainText.Extract(paragraph.Inlines);
        }

        return sourceCell.Markdown;
    }
}

/// <summary>
/// Native projection for a markdown block without a specialized projection.
/// </summary>
public sealed class MarkdownNativeOtherBlock : MarkdownNativeBlock {
    internal MarkdownNativeOtherBlock(IMarkdownBlock block, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Other, block, syntaxNode) {
        Markdown = block.RenderMarkdown();
    }

    /// <summary>Markdown representation of the source block.</summary>
    public string Markdown { get; }
}
