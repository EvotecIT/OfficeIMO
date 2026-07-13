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

    /// <summary>Enumerates source-backed fields owned by this block in source order.</summary>
    public IEnumerable<MarkdownNativeBlockSourceField> EnumerateSourceFields() =>
        MarkdownNativeDocument.EnumerateBlockSourceFields(this);

    /// <summary>Enumerates source-backed fields with the supplied field name owned by this block in source order.</summary>
    public IEnumerable<MarkdownNativeBlockSourceField> EnumerateSourceFields(string name) {
        if (string.IsNullOrWhiteSpace(name)) {
            yield break;
        }

        foreach (var field in EnumerateSourceFields()) {
            if (string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase)) {
                yield return field;
            }
        }
    }

    /// <summary>
    /// Finds the first source-backed field with the supplied name, optionally constrained to a repeated-field occurrence index.
    /// </summary>
    public MarkdownNativeBlockSourceField? FindSourceField(string name, int index = -1) {
        if (string.IsNullOrWhiteSpace(name)) {
            return null;
        }

        foreach (var field in EnumerateSourceFields()) {
            if (!string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            if (index < 0 || field.Index == index) {
                return field;
            }
        }

        return null;
    }

    /// <summary>Creates a UI-safe snapshot of this block without parser object references.</summary>
    public MarkdownNativeBlockSnapshot ToSnapshot() => MarkdownNativeSnapshotFactory.FromBlock(this);
}

/// <summary>
/// Native projection for a paragraph block.
/// </summary>
public sealed class MarkdownNativeParagraphBlock : MarkdownNativeBlock {
    internal MarkdownNativeParagraphBlock(ParagraphBlock paragraph, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Paragraph, paragraph, syntaxNode) {
        Paragraph = paragraph;
        Inlines = paragraph.Inlines;
        InlineRuns = MarkdownNativeInlineProjection.FromInlineContainer(syntaxNode);
        Text = InlinePlainText.Extract(paragraph.Inlines);
        TextSourceSpan = GetParagraphTextSourceSpan(syntaxNode) ?? paragraph.SourceSpan;
    }

    /// <summary>Source paragraph block.</summary>
    public ParagraphBlock Paragraph { get; }

    /// <summary>Plain-text paragraph content.</summary>
    public string Text { get; }

    /// <summary>Source span for the paragraph payload when available.</summary>
    public MarkdownSourceSpan? TextSourceSpan { get; }

    /// <summary>Structured inline nodes.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>AST-backed native inline projection with source spans.</summary>
    public IReadOnlyList<MarkdownNativeInline> InlineRuns { get; }

    private static MarkdownSourceSpan? GetParagraphTextSourceSpan(MarkdownSyntaxNode syntaxNode) {
        if (syntaxNode.Children.Count == 0) {
            return syntaxNode.SourceSpan;
        }

        var contentChildren = syntaxNode.Children
            .Where(static child => child.Kind != MarkdownSyntaxKind.GenericAttributeBlock)
            .ToArray();
        return MarkdownBlockSyntaxBuilder.GetAggregateSpan(contentChildren) ?? syntaxNode.SourceSpan;
    }
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
        Attributes = code.Attributes.Attributes;
        Classes = code.Attributes.Classes;
        ElementId = code.Attributes.ElementId;
        Title = code.Attributes.GetAttribute("title");
        OpeningFence = code.IsFenced ? GetChildLiteral(syntaxNode, MarkdownSyntaxKind.CodeFenceOpening) ?? new string(code.FenceChar, code.FenceLength) : null;
        ClosingFence = code.IsFenced && code.HasClosingFence ? GetChildLiteral(syntaxNode, MarkdownSyntaxKind.CodeFenceClosing) ?? new string(code.FenceChar, code.ClosingFenceLength) : null;
        OpeningFenceSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.CodeFenceOpening) ?? code.OpeningFenceSourceSpan;
        InfoStringSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.CodeFenceInfo) ?? code.InfoStringSourceSpan;
        AttributeSourceSpan = MarkdownNativeFenceInfoSourceSpans.GetAttributeSourceSpan(FenceInfo, InfoStringSourceSpan);
        ContentSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.CodeContent) ?? code.ContentSourceSpan;
        ClosingFenceSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.CodeFenceClosing) ?? code.ClosingFenceSourceSpan;
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

    /// <summary>Opening fence marker text, preserving the parsed marker character and length when source-backed.</summary>
    public string? OpeningFence { get; }

    /// <summary>Closing fence marker text, preserving the parsed marker character and length when source-backed.</summary>
    public string? ClosingFence { get; }

    /// <summary>Source span for the opening fence marker when the block was parsed from a fenced source block.</summary>
    public MarkdownSourceSpan? OpeningFenceSourceSpan { get; }

    /// <summary>Source span for the fenced-code info string when the block was parsed from a fenced source block.</summary>
    public MarkdownSourceSpan? InfoStringSourceSpan { get; }

    /// <summary>Source span for explicit generic attributes inside the fenced-code info string.</summary>
    public MarkdownSourceSpan? AttributeSourceSpan { get; }

    /// <summary>Source span for the code payload when available.</summary>
    public MarkdownSourceSpan? ContentSourceSpan { get; }

    /// <summary>Source span for the closing fence marker when the block was parsed from a closed fenced source block.</summary>
    public MarkdownSourceSpan? ClosingFenceSourceSpan { get; }

    private static MarkdownSourceSpan? GetChildSpan(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) =>
        syntaxNode?.Children.FirstOrDefault(child => child.Kind == kind)?.SourceSpan;

    private static string? GetChildLiteral(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) =>
        syntaxNode?.Children.FirstOrDefault(child => child.Kind == kind)?.Literal;
}

/// <summary>
/// Native projection for a CommonMark thematic break / horizontal rule.
/// </summary>
public sealed class MarkdownNativeThematicBreakBlock : MarkdownNativeBlock {
    internal MarkdownNativeThematicBreakBlock(HorizontalRuleBlock horizontalRule, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.ThematicBreak, horizontalRule, syntaxNode) {
        HorizontalRule = horizontalRule;
        Marker = syntaxNode.Literal ?? ((IMarkdownBlock)horizontalRule).RenderMarkdown();
        MarkerText = horizontalRule.MarkerText ?? Marker;
        MarkerSourceSpan = horizontalRule.MarkerSourceSpan ?? syntaxNode.SourceSpan;
    }

    /// <summary>Source horizontal rule block.</summary>
    public HorizontalRuleBlock HorizontalRule { get; }

    /// <summary>Normalized markdown marker used for semantic rendering.</summary>
    public string Marker { get; }

    /// <summary>Exact thematic-break marker text when parsed from markdown.</summary>
    public string MarkerText { get; }

    /// <summary>Source span for the thematic-break marker token when available.</summary>
    public MarkdownSourceSpan? MarkerSourceSpan { get; }
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
        Payload = MarkdownNativeVisualPayload.Create(visual);
        OpeningFence = visual.IsFenced ? GetChildLiteral(syntaxNode, MarkdownSyntaxKind.CodeFenceOpening) ?? new string(visual.FenceChar, visual.FenceLength) : null;
        ClosingFence = visual.IsFenced && visual.HasClosingFence ? GetChildLiteral(syntaxNode, MarkdownSyntaxKind.CodeFenceClosing) ?? new string(visual.FenceChar, visual.ClosingFenceLength) : null;
        OpeningFenceSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.CodeFenceOpening) ?? visual.OpeningFenceSourceSpan;
        InfoStringSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.CodeFenceInfo) ?? visual.InfoStringSourceSpan;
        AttributeSourceSpan = MarkdownNativeFenceInfoSourceSpans.GetAttributeSourceSpan(FenceInfo, InfoStringSourceSpan);
        ContentSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.CodeContent) ?? visual.ContentSourceSpan;
        ClosingFenceSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.CodeFenceClosing) ?? visual.ClosingFenceSourceSpan;
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

    /// <summary>Dependency-free typed payload hints for visual UI hosts.</summary>
    public MarkdownNativeVisualPayload Payload { get; }

    /// <summary>Opening fence marker text, preserving the parsed marker character and length when source-backed.</summary>
    public string? OpeningFence { get; }

    /// <summary>Closing fence marker text, preserving the parsed marker character and length when source-backed.</summary>
    public string? ClosingFence { get; }

    /// <summary>Source span for the opening fence marker when the block was parsed from a fenced source block.</summary>
    public MarkdownSourceSpan? OpeningFenceSourceSpan { get; }

    /// <summary>Source span for the fenced-block info string when the block was parsed from a fenced source block.</summary>
    public MarkdownSourceSpan? InfoStringSourceSpan { get; }

    /// <summary>Source span for explicit generic attributes inside the fenced-block info string.</summary>
    public MarkdownSourceSpan? AttributeSourceSpan { get; }

    /// <summary>Source span for the fenced payload when available.</summary>
    public MarkdownSourceSpan? ContentSourceSpan { get; }

    /// <summary>Source span for the closing fence marker when the block was parsed from a closed fenced source block.</summary>
    public MarkdownSourceSpan? ClosingFenceSourceSpan { get; }

    private static MarkdownSourceSpan? GetChildSpan(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) =>
        syntaxNode?.Children.FirstOrDefault(child => child.Kind == kind)?.SourceSpan;

    private static string? GetChildLiteral(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) =>
        syntaxNode?.Children.FirstOrDefault(child => child.Kind == kind)?.Literal;
}

/// <summary>
/// Native projection for a markdown table.
/// </summary>
public sealed class MarkdownNativeTableBlock : MarkdownNativeBlock {
    internal MarkdownNativeTableBlock(
        TableBlock table,
        MarkdownSyntaxNode syntaxNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics)
        : base(MarkdownNativeBlockKind.Table, table, syntaxNode) {
        Table = table;
        AlignmentRowSourceSpan = syntaxNode.Children.FirstOrDefault(static child => child.Kind == MarkdownSyntaxKind.TableAlignmentRow)?.SourceSpan;
        AlignmentCells = BuildAlignmentCells(table);
        Pipes = BuildPipes(table);
        HeaderCells = BuildHeaderCells(table, syntaxNode, diagnostics);
        Rows = BuildRows(table, syntaxNode, diagnostics);
        HeaderRow = BuildHeaderRow(table, syntaxNode, HeaderCells);
        BodyRows = BuildNativeRows(table, syntaxNode, Rows);
    }

    /// <summary>Source table block.</summary>
    public TableBlock Table { get; }

    /// <summary>Source span for the GFM table alignment/separator row when present.</summary>
    public MarkdownSourceSpan? AlignmentRowSourceSpan { get; }

    /// <summary>Per-column alignment markers from the GFM table alignment/separator row.</summary>
    public IReadOnlyList<MarkdownNativeTableAlignmentCell> AlignmentCells { get; }

    /// <summary>Pipe delimiter tokens in document order.</summary>
    public IReadOnlyList<MarkdownNativeTablePipe> Pipes { get; }

    /// <summary>Header cells in document order.</summary>
    public IReadOnlyList<MarkdownNativeTableCell> HeaderCells { get; }

    /// <summary>Header row projection when the table has a header row.</summary>
    public MarkdownNativeTableRow? HeaderRow { get; }

    /// <summary>Body rows and cells in document order.</summary>
    public IReadOnlyList<IReadOnlyList<MarkdownNativeTableCell>> Rows { get; }

    /// <summary>Body row projections in document order.</summary>
    public IReadOnlyList<MarkdownNativeTableRow> BodyRows { get; }

    private static IReadOnlyList<MarkdownNativeTableAlignmentCell> BuildAlignmentCells(TableBlock table) {
        if (table.AlignmentCellSources.Count == 0) {
            return Array.Empty<MarkdownNativeTableAlignmentCell>();
        }

        var cells = new List<MarkdownNativeTableAlignmentCell>(table.AlignmentCellSources.Count);
        for (var columnIndex = 0; columnIndex < table.AlignmentCellSources.Count; columnIndex++) {
            var source = table.AlignmentCellSources[columnIndex];
            var alignment = table.Alignments != null && columnIndex < table.Alignments.Count
                ? table.Alignments[columnIndex]
                : ColumnAlignment.None;
            cells.Add(new MarkdownNativeTableAlignmentCell(columnIndex, alignment, source.Markdown, source.SourceSpan));
        }

        return cells;
    }

    private static IReadOnlyList<MarkdownNativeTablePipe> BuildPipes(TableBlock table) {
        if (table.PipeSources.Count == 0) {
            return Array.Empty<MarkdownNativeTablePipe>();
        }

        var pipes = new List<MarkdownNativeTablePipe>(table.PipeSources.Count);
        for (var index = 0; index < table.PipeSources.Count; index++) {
            var source = table.PipeSources[index];
            pipes.Add(new MarkdownNativeTablePipe(source.RowIndex, source.ColumnIndex, source.SourceSpan));
        }

        return pipes;
    }

    private static IReadOnlyList<MarkdownNativeTableCell> BuildHeaderCells(
        TableBlock table,
        MarkdownSyntaxNode syntaxNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics) {
        if (table.Headers.Count == 0) {
            return Array.Empty<MarkdownNativeTableCell>();
        }

        var headerNode = syntaxNode.Children.FirstOrDefault(static child => child.Kind == MarkdownSyntaxKind.TableHeader);
        return BuildCells(table.Headers, table.HeaderCells, table.Alignments, headerNode, diagnostics, isHeader: true, rowIndex: -1);
    }

    private static IReadOnlyList<IReadOnlyList<MarkdownNativeTableCell>> BuildRows(
        TableBlock table,
        MarkdownSyntaxNode syntaxNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics) {
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
            rows.Add(BuildCells(rawCells, structuredCells, table.Alignments, rowNode, diagnostics, isHeader: false, rowIndex: rowIndex));
        }

        return rows;
    }

    private static MarkdownNativeTableRow? BuildHeaderRow(
        TableBlock table,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<MarkdownNativeTableCell> headerCells) {
        if (table.Headers.Count == 0) {
            return null;
        }

        var headerNode = syntaxNode.Children.FirstOrDefault(static child => child.Kind == MarkdownSyntaxKind.TableHeader);
        return new MarkdownNativeTableRow(table.HeaderRow, headerNode, isHeader: true, rowIndex: -1, headerCells);
    }

    private static IReadOnlyList<MarkdownNativeTableRow> BuildNativeRows(
        TableBlock table,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<IReadOnlyList<MarkdownNativeTableCell>> rows) {
        if (rows == null || rows.Count == 0) {
            return Array.Empty<MarkdownNativeTableRow>();
        }

        var rowNodes = syntaxNode.Children.Where(static child => child.Kind == MarkdownSyntaxKind.TableRow).ToArray();
        var sourceRows = table.BodyRows;
        var nativeRows = new MarkdownNativeTableRow[rows.Count];
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var sourceRow = rowIndex < sourceRows.Count ? sourceRows[rowIndex] : null;
            var rowNode = rowIndex < rowNodes.Length ? rowNodes[rowIndex] : null;
            nativeRows[rowIndex] = new MarkdownNativeTableRow(sourceRow, rowNode, isHeader: false, rowIndex, rows[rowIndex]);
        }

        return nativeRows;
    }

    private static IReadOnlyList<MarkdownNativeTableCell> BuildCells(
        IReadOnlyList<string> rawCells,
        IReadOnlyList<TableCell> structuredCells,
        IReadOnlyList<ColumnAlignment> columnAlignments,
        MarkdownSyntaxNode? rowNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics,
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
            var children = cellNode != null
                ? MarkdownNativeProjectionFactory.CreateChildren(cellNode, diagnostics)
                : Array.Empty<MarkdownNativeBlock>();
            cells.Add(new MarkdownNativeTableCell(raw, cell, cellNode, isHeader, rowIndex, columnIndex, alignment, children));
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
/// Native projection for a table row.
/// </summary>
public sealed class MarkdownNativeTableRow {
    internal MarkdownNativeTableRow(
        TableRow? sourceRow,
        MarkdownSyntaxNode? syntaxNode,
        bool isHeader,
        int rowIndex,
        IReadOnlyList<MarkdownNativeTableCell> cells) {
        SourceRow = sourceRow;
        SyntaxNode = syntaxNode;
        SourceSpan = syntaxNode?.SourceSpan ?? sourceRow?.SourceSpan;
        IsHeader = isHeader;
        RowIndex = rowIndex;
        Cells = cells ?? Array.Empty<MarkdownNativeTableCell>();
        Markdown = syntaxNode?.Literal ?? string.Join(" | ", Cells.Select(static cell => cell.Markdown));
    }

    /// <summary>Source table row when structured row data is available.</summary>
    public TableRow? SourceRow { get; }

    /// <summary>Syntax node that produced this row when available.</summary>
    public MarkdownSyntaxNode? SyntaxNode { get; }

    /// <summary>Source span in the normalized markdown text when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Whether this row is the table header row.</summary>
    public bool IsHeader { get; }

    /// <summary>Zero-based body row index, or <c>-1</c> for the header row.</summary>
    public int RowIndex { get; }

    /// <summary>Cells in document column order.</summary>
    public IReadOnlyList<MarkdownNativeTableCell> Cells { get; }

    /// <summary>Markdown representation of the row payload.</summary>
    public string Markdown { get; }
}

/// <summary>
/// Native projection for a source-backed pipe delimiter in a table row.
/// </summary>
public sealed class MarkdownNativeTablePipe {
    internal MarkdownNativeTablePipe(int rowIndex, int columnIndex, MarkdownSourceSpan sourceSpan) {
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        SourceSpan = sourceSpan;
    }

    /// <summary>
    /// Row index for the delimiter: <c>-1</c> for the header row, <c>-2</c> for the alignment row, or zero-based body row index.
    /// </summary>
    public int RowIndex { get; }

    /// <summary>Zero-based delimiter occurrence within the row.</summary>
    public int ColumnIndex { get; }

    /// <summary>Source span for this pipe delimiter in the normalized markdown text.</summary>
    public MarkdownSourceSpan SourceSpan { get; }
}

/// <summary>
/// Native projection for a table alignment/separator row cell.
/// </summary>
public sealed class MarkdownNativeTableAlignmentCell {
    internal MarkdownNativeTableAlignmentCell(
        int columnIndex,
        ColumnAlignment alignment,
        string markdown,
        MarkdownSourceSpan sourceSpan) {
        ColumnIndex = columnIndex;
        Alignment = alignment;
        Markdown = markdown ?? string.Empty;
        SourceSpan = sourceSpan;
    }

    /// <summary>Zero-based column index.</summary>
    public int ColumnIndex { get; }

    /// <summary>Semantic alignment represented by this alignment cell.</summary>
    public ColumnAlignment Alignment { get; }

    /// <summary>Original alignment marker text such as <c>:---</c>, <c>---:</c>, or <c>:---:</c>.</summary>
    public string Markdown { get; }

    /// <summary>Source span for this alignment marker in the normalized markdown text.</summary>
    public MarkdownSourceSpan SourceSpan { get; }
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
        ColumnAlignment alignment,
        IReadOnlyList<MarkdownNativeBlock> children) {
        RawText = rawText ?? string.Empty;
        SourceCell = sourceCell;
        SyntaxNode = syntaxNode;
        SourceSpan = syntaxNode?.SourceSpan ?? sourceCell?.SourceSpan;
        IsHeader = isHeader;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Text = ExtractText(sourceCell, RawText);
        Markdown = sourceCell?.Markdown ?? RawText;
        Blocks = sourceCell != null ? sourceCell.ChildBlocks : Array.Empty<IMarkdownBlock>();
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
        Alignment = alignment;
        InlineRuns = MarkdownNativeInlineProjection.FromTableCellDirectContent(syntaxNode);
    }

    /// <summary>Raw cell text from the table source.</summary>
    public string RawText { get; }

    /// <summary>Plain-text cell content.</summary>
    public string Text { get; }

    /// <summary>Markdown representation of the cell content.</summary>
    public string Markdown { get; }

    /// <summary>Structured child blocks in the cell.</summary>
    public IReadOnlyList<IMarkdownBlock> Blocks { get; }

    /// <summary>Native child blocks projected from structured cell content.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }

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

    /// <summary>AST-backed native inline projection for the cell content when available.</summary>
    public IReadOnlyList<MarkdownNativeInline> InlineRuns { get; }

    private static string ExtractText(TableCell? sourceCell, string rawText) {
        if (sourceCell == null || sourceCell.ChildBlocks.Count == 0) {
            return rawText ?? string.Empty;
        }

        if (sourceCell.ChildBlocks.Count == 1 && sourceCell.ChildBlocks[0] is ParagraphBlock paragraph) {
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
