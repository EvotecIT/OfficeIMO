namespace OfficeIMO.Markdown;

/// <summary>
/// Native, AST-backed projection of a parsed markdown document for UI hosts that need structured blocks and source spans.
/// </summary>
public sealed partial class MarkdownNativeDocument {
    private MarkdownNativeDocument(
        MarkdownParseResult parseResult,
        string sourceMarkdown,
        MarkdownNativeDocumentSourceKind sourceKind,
        IReadOnlyList<MarkdownNativeBlock> blocks,
        IReadOnlyList<MarkdownNativeSourceTrivia> sourceTrivia,
        IReadOnlyList<MarkdownNativeDiagnostic> diagnostics) {
        ParseResult = parseResult ?? throw new ArgumentNullException(nameof(parseResult));
        Document = parseResult.Document;
        SyntaxTree = parseResult.SyntaxTree;
        FinalSyntaxTree = parseResult.FinalSyntaxTree;
        TransformDiagnostics = parseResult.TransformDiagnostics;
        SourceMarkdown = sourceMarkdown ?? string.Empty;
        SourceKind = sourceKind;
        Blocks = blocks ?? Array.Empty<MarkdownNativeBlock>();
        SourceTrivia = sourceTrivia ?? Array.Empty<MarkdownNativeSourceTrivia>();
        Diagnostics = diagnostics ?? Array.Empty<MarkdownNativeDiagnostic>();
    }

    /// <summary>Underlying parse result, including original/final syntax trees and diagnostics.</summary>
    public MarkdownParseResult ParseResult { get; }

    /// <summary>Parsed OfficeIMO markdown document.</summary>
    public MarkdownDoc Document { get; }

    /// <summary>Original syntax tree produced before document transforms were applied.</summary>
    public MarkdownSyntaxNode SyntaxTree { get; }

    /// <summary>Final syntax tree aligned with <see cref="Document"/>.</summary>
    public MarkdownSyntaxNode FinalSyntaxTree { get; }

    /// <summary>Document-transform diagnostics captured during parsing.</summary>
    public IReadOnlyList<MarkdownDocumentTransformDiagnostic> TransformDiagnostics { get; }

    /// <summary>Effective reference-style link definitions collected during parsing.</summary>
    public IReadOnlyList<MarkdownReferenceLinkDefinition> ReferenceLinkDefinitions => ParseResult.ReferenceLinkDefinitions;

    /// <summary>Markdown source text whose source spans back this projection.</summary>
    public string SourceMarkdown { get; }

    /// <summary>Identifies whether <see cref="SourceMarkdown"/> is direct reader input or renderer-preprocessed markdown.</summary>
    public MarkdownNativeDocumentSourceKind SourceKind { get; }

    /// <summary>Top-level native block projection in document order.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Blocks { get; }

    /// <summary>Document-level source trivia such as blank lines, in source order.</summary>
    public IReadOnlyList<MarkdownNativeSourceTrivia> SourceTrivia { get; }

    /// <summary>Projection diagnostics including transform notices and unsupported block fallbacks.</summary>
    public IReadOnlyList<MarkdownNativeDiagnostic> Diagnostics { get; }

    /// <summary>
    /// Parses markdown into the typed object model, syntax tree, diagnostics, and native block projection.
    /// </summary>
    public static MarkdownNativeDocument Parse(string markdown, MarkdownReaderOptions? options = null) {
        markdown ??= string.Empty;
        var parseResult = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);
        return FromParseResult(parseResult, sourceMarkdown: null, MarkdownNativeDocumentSourceKind.ReaderInput);
    }

    /// <summary>
    /// Builds a native projection from an existing syntax-backed parse result.
    /// </summary>
    public static MarkdownNativeDocument FromParseResult(
        MarkdownParseResult parseResult,
        string? sourceMarkdown = null,
        MarkdownNativeDocumentSourceKind sourceKind = MarkdownNativeDocumentSourceKind.ReaderInput) {
        if (parseResult == null) {
            throw new ArgumentNullException(nameof(parseResult));
        }

        var blocks = new List<MarkdownNativeBlock>();
        var diagnostics = new List<MarkdownNativeDiagnostic>();
        for (var i = 0; i < parseResult.TransformDiagnostics.Count; i++) {
            diagnostics.Add(MarkdownNativeDiagnostic.FromTransform(parseResult.TransformDiagnostics[i]));
        }

        var children = parseResult.FinalSyntaxTree.Children;
        for (var i = 0; i < children.Count; i++) {
            var block = MarkdownNativeProjectionFactory.Create(children[i], diagnostics);
            if (block != null) {
                blocks.Add(block);
            }
        }

        var nativeSourceMarkdown = sourceMarkdown ?? parseResult.SourceMarkdown;
        return new MarkdownNativeDocument(
            parseResult,
            nativeSourceMarkdown,
            sourceKind,
            blocks,
            CreateSourceTrivia(nativeSourceMarkdown),
            diagnostics);
    }

    /// <summary>Finds the first native block whose source span contains the supplied 1-based line.</summary>
    public MarkdownNativeBlock? FindBlockAtLine(int lineNumber) {
        for (var i = 0; i < Blocks.Count; i++) {
            var match = FindBlockAtLine(Blocks[i], lineNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Finds the first native block whose source span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeBlock? FindBlockAtPosition(int lineNumber, int columnNumber) {
        for (var i = 0; i < Blocks.Count; i++) {
            var match = FindBlockAtPosition(Blocks[i], lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Finds a native block by deterministic id.</summary>
    public MarkdownNativeBlock? FindBlockById(string id) {
        if (string.IsNullOrWhiteSpace(id)) {
            return null;
        }

        for (var i = 0; i < Blocks.Count; i++) {
            var match = FindBlockById(Blocks[i], id);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Finds the first native inline whose source span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeInline? FindInlineAtPosition(int lineNumber, int columnNumber) {
        foreach (var inline in EnumerateInlines()) {
            var match = FindInlineAtPosition(inline, lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Finds a native inline by deterministic id.</summary>
    public MarkdownNativeInline? FindInlineById(string id) {
        if (string.IsNullOrWhiteSpace(id)) {
            return null;
        }

        foreach (var inline in EnumerateInlines()) {
            var match = FindInlineById(inline, id);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Returns the top-level-to-target block path for a block id.</summary>
    public IReadOnlyList<MarkdownNativeBlock> GetBlockPath(string id) {
        if (string.IsNullOrWhiteSpace(id)) {
            return Array.Empty<MarkdownNativeBlock>();
        }

        var path = new List<MarkdownNativeBlock>();
        for (var i = 0; i < Blocks.Count; i++) {
            if (TryBuildBlockPath(Blocks[i], id, path)) {
                return path;
            }
        }

        return Array.Empty<MarkdownNativeBlock>();
    }

    /// <summary>Enumerates top-level native blocks of the requested projection type.</summary>
    public IEnumerable<TBlock> BlocksOfType<TBlock>() where TBlock : MarkdownNativeBlock {
        for (var i = 0; i < Blocks.Count; i++) {
            if (Blocks[i] is TBlock block) {
                yield return block;
            }
        }
    }

    /// <summary>Enumerates all native blocks in document order, including nested blocks.</summary>
    public IEnumerable<MarkdownNativeBlock> DescendantBlocksAndSelf() {
        for (var i = 0; i < Blocks.Count; i++) {
            foreach (var block in DescendantBlocksAndSelf(Blocks[i])) {
                yield return block;
            }
        }
    }

    /// <summary>Enumerates all native inline runs in document order.</summary>
    public IEnumerable<MarkdownNativeInline> EnumerateInlines() {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (var i = 0; i < Blocks.Count; i++) {
            foreach (var inline in EnumerateInlines(Blocks[i])) {
                if (seen.Add(inline.Id)) {
                    yield return inline;
                }
            }
        }
    }

    /// <summary>Creates a UI-safe snapshot of this document without parser object references.</summary>
    public MarkdownNativeDocumentSnapshot ToSnapshot() => MarkdownNativeSnapshotFactory.FromDocument(this);

    /// <summary>
    /// Emits the original markdown when this native document was parsed with trivia preservation and no transforms changed it.
    /// </summary>
    public MarkdownRoundtripResult WriteUnchanged() => MarkdownRoundtripWriter.WriteUnchanged(ParseResult);

    /// <summary>
    /// Applies one native source edit while preserving original markdown around it when the edit can be remapped safely.
    /// </summary>
    public MarkdownRoundtripResult WriteWithSourceEdit(MarkdownNativeSourceEdit edit) =>
        MarkdownRoundtripWriter.WriteWithSourceEdit(ParseResult, edit);

    /// <summary>
    /// Applies native source edits while preserving original markdown around them when every edit can be remapped safely.
    /// </summary>
    public MarkdownRoundtripResult WriteWithSourceEdits(IEnumerable<MarkdownNativeSourceEdit> edits) =>
        MarkdownRoundtripWriter.WriteWithSourceEdits(ParseResult, edits);

    /// <summary>Creates a non-mutating source edit that replaces a source span.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownSourceSpan sourceSpan, string replacementMarkdown) {
        if (!TryResolveOffsets(sourceSpan, out var startOffset, out var endOffsetInclusive)) {
            throw new InvalidOperationException("The supplied source span cannot be mapped to offsets in this native document source.");
        }

        return new MarkdownNativeSourceEdit(sourceSpan, startOffset, endOffsetInclusive, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces a native block.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeBlock block, string replacementMarkdown) {
        if (block == null) {
            throw new ArgumentNullException(nameof(block));
        }

        if (!block.SourceSpan.HasValue) {
            throw new InvalidOperationException("The native block does not have a source span.");
        }

        return CreateReplaceEdit(block.SourceSpan.Value, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces a native inline.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeInline inline, string replacementMarkdown) {
        if (inline == null) {
            throw new ArgumentNullException(nameof(inline));
        }

        if (!inline.SourceSpan.HasValue) {
            throw new InvalidOperationException("The native inline does not have a source span.");
        }

        return CreateReplaceEdit(inline.SourceSpan.Value, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces the source-backed content span of a native list item.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeListItem listItem, string replacementMarkdown) {
        if (listItem == null) {
            throw new ArgumentNullException(nameof(listItem));
        }

        if (!listItem.ContentSourceSpan.HasValue) {
            throw new InvalidOperationException("The native list item does not have a source-backed content span.");
        }

        return CreateReplaceEdit(listItem.ContentSourceSpan.Value, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces a reference-style link definition.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownReferenceLinkDefinition referenceDefinition, string replacementMarkdown) {
        if (referenceDefinition == null) {
            throw new ArgumentNullException(nameof(referenceDefinition));
        }

        if (!referenceDefinition.SourceSpan.HasValue) {
            throw new InvalidOperationException("The reference definition does not have a source span.");
        }

        return CreateReplaceEdit(referenceDefinition.SourceSpan.Value, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces a native table cell.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeTableCell tableCell, string replacementMarkdown) {
        if (tableCell == null) {
            throw new ArgumentNullException(nameof(tableCell));
        }

        if (!tableCell.SourceSpan.HasValue) {
            throw new InvalidOperationException("The native table cell does not have a source span.");
        }

        return CreateReplaceEdit(tableCell.SourceSpan.Value, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces a native definition-list group.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeDefinitionListGroup definitionGroup, string replacementMarkdown) {
        if (definitionGroup == null) {
            throw new ArgumentNullException(nameof(definitionGroup));
        }

        if (!definitionGroup.SourceSpan.HasValue) {
            throw new InvalidOperationException("The native definition-list group does not have a source span.");
        }

        return CreateReplaceEdit(definitionGroup.SourceSpan.Value, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces a native definition-list term.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeDefinitionListTerm definitionTerm, string replacementMarkdown) {
        if (definitionTerm == null) {
            throw new ArgumentNullException(nameof(definitionTerm));
        }

        if (!definitionTerm.SourceSpan.HasValue) {
            throw new InvalidOperationException("The native definition-list term does not have a source span.");
        }

        return CreateReplaceEdit(definitionTerm.SourceSpan.Value, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces a native definition-list definition body.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeDefinitionListDefinition definition, string replacementMarkdown) {
        if (definition == null) {
            throw new ArgumentNullException(nameof(definition));
        }

        if (!definition.SourceSpan.HasValue) {
            throw new InvalidOperationException("The native definition-list definition does not have a source span.");
        }

        return CreateReplaceEdit(
            definition.SourceSpan.Value,
            FormatDefinitionListDefinitionReplacement(definition, replacementMarkdown));
    }

    /// <summary>Creates a non-mutating source edit that replaces source-backed inline metadata.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeInlineMetadata metadata, string replacementMarkdown) {
        if (metadata == null) {
            throw new ArgumentNullException(nameof(metadata));
        }

        if (!metadata.SourceSpan.HasValue) {
            throw new InvalidOperationException("The native inline metadata does not have a source span.");
        }

        return CreateReplaceEdit(metadata.SourceSpan.Value, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces document-level source trivia.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeSourceTrivia trivia, string replacementMarkdown) {
        if (trivia == null) {
            throw new ArgumentNullException(nameof(trivia));
        }

        return CreateReplaceEdit(trivia.SourceSpan, replacementMarkdown);
    }

    private static MarkdownNativeBlock? FindBlockAtLine(MarkdownNativeBlock block, int lineNumber) {
        switch (block) {
            case MarkdownNativeQuoteBlock quote:
                return FindChildBlockAtLine(quote.Children, lineNumber) ?? (quote.ContainsLine(lineNumber) ? quote : null);
            case MarkdownNativeCalloutBlock callout:
                return FindChildBlockAtLine(callout.Children, lineNumber) ?? (callout.ContainsLine(lineNumber) ? callout : null);
            case MarkdownNativeDetailsBlock details:
                return FindChildBlockAtLine(details.Children, lineNumber) ?? (details.ContainsLine(lineNumber) ? details : null);
            case MarkdownNativeDefinitionListBlock definitionList:
                return FindChildBlockAtLine(definitionList.Children, lineNumber) ?? (definitionList.ContainsLine(lineNumber) ? definitionList : null);
            case MarkdownNativeFootnoteDefinitionBlock footnote:
                return FindChildBlockAtLine(footnote.Children, lineNumber) ?? (footnote.ContainsLine(lineNumber) ? footnote : null);
            case MarkdownNativeListBlock list:
                for (var i = 0; i < list.Items.Count; i++) {
                    var itemMatch = FindChildBlockAtLine(list.Items[i].Children, lineNumber);
                    if (itemMatch != null) {
                        return itemMatch;
                    }
                }

                return list.ContainsLine(lineNumber) ? list : null;
            case MarkdownNativeTableBlock table:
                return FindChildBlockAtLine(EnumerateTableChildren(table), lineNumber) ?? (table.ContainsLine(lineNumber) ? table : null);
            default:
                return block.ContainsLine(lineNumber) ? block : null;
        }
    }

    private static MarkdownNativeBlock? FindBlockAtPosition(MarkdownNativeBlock block, int lineNumber, int columnNumber) {
        switch (block) {
            case MarkdownNativeQuoteBlock quote:
                return FindChildBlockAtPosition(quote.Children, lineNumber, columnNumber) ?? (ContainsPosition(quote, lineNumber, columnNumber) ? quote : null);
            case MarkdownNativeCalloutBlock callout:
                return FindChildBlockAtPosition(callout.Children, lineNumber, columnNumber) ?? (ContainsPosition(callout, lineNumber, columnNumber) ? callout : null);
            case MarkdownNativeDetailsBlock details:
                return FindChildBlockAtPosition(details.Children, lineNumber, columnNumber) ?? (ContainsPosition(details, lineNumber, columnNumber) ? details : null);
            case MarkdownNativeDefinitionListBlock definitionList:
                return FindChildBlockAtPosition(definitionList.Children, lineNumber, columnNumber) ?? (ContainsPosition(definitionList, lineNumber, columnNumber) ? definitionList : null);
            case MarkdownNativeFootnoteDefinitionBlock footnote:
                return FindChildBlockAtPosition(footnote.Children, lineNumber, columnNumber) ?? (ContainsPosition(footnote, lineNumber, columnNumber) ? footnote : null);
            case MarkdownNativeListBlock list:
                for (var i = 0; i < list.Items.Count; i++) {
                    var itemMatch = FindChildBlockAtPosition(list.Items[i].Children, lineNumber, columnNumber);
                    if (itemMatch != null) {
                        return itemMatch;
                    }
                }

                return ContainsPosition(list, lineNumber, columnNumber) ? list : null;
            case MarkdownNativeTableBlock table:
                return FindChildBlockAtPosition(EnumerateTableChildren(table), lineNumber, columnNumber) ?? (ContainsPosition(table, lineNumber, columnNumber) ? table : null);
            default:
                return ContainsPosition(block, lineNumber, columnNumber) ? block : null;
        }
    }

    private static bool ContainsPosition(MarkdownNativeBlock block, int lineNumber, int columnNumber) =>
        block.SourceSpan.HasValue && block.SourceSpan.Value.ContainsPosition(lineNumber, columnNumber);

    private static MarkdownNativeBlock? FindBlockById(MarkdownNativeBlock block, string id) {
        if (string.Equals(block.Id, id, StringComparison.Ordinal)) {
            return block;
        }

        foreach (var child in GetChildBlocks(block)) {
            var match = FindBlockById(child, id);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static bool TryBuildBlockPath(MarkdownNativeBlock block, string id, List<MarkdownNativeBlock> path) {
        path.Add(block);
        if (string.Equals(block.Id, id, StringComparison.Ordinal)) {
            return true;
        }

        foreach (var child in GetChildBlocks(block)) {
            if (TryBuildBlockPath(child, id, path)) {
                return true;
            }
        }

        path.RemoveAt(path.Count - 1);
        return false;
    }

    private static MarkdownNativeBlock? FindChildBlockAtLine(IReadOnlyList<MarkdownNativeBlock> children, int lineNumber) {
        for (var i = 0; i < children.Count; i++) {
            var match = FindBlockAtLine(children[i], lineNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static MarkdownNativeBlock? FindChildBlockAtPosition(IReadOnlyList<MarkdownNativeBlock> children, int lineNumber, int columnNumber) {
        for (var i = 0; i < children.Count; i++) {
            var match = FindBlockAtPosition(children[i], lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static IEnumerable<MarkdownNativeBlock> DescendantBlocksAndSelf(MarkdownNativeBlock block) {
        yield return block;
        foreach (var child in GetChildBlocks(block)) {
            foreach (var descendant in DescendantBlocksAndSelf(child)) {
                yield return descendant;
            }
        }
    }

    private static IEnumerable<MarkdownNativeBlock> GetChildBlocks(MarkdownNativeBlock block) {
        switch (block) {
            case MarkdownNativeQuoteBlock quote:
                return quote.Children;
            case MarkdownNativeCalloutBlock callout:
                return callout.Children;
            case MarkdownNativeDetailsBlock details:
                return details.Children;
            case MarkdownNativeDefinitionListBlock definitionList:
                return definitionList.Children;
            case MarkdownNativeFootnoteDefinitionBlock footnote:
                return footnote.Children;
            case MarkdownNativeListBlock list:
                return EnumerateListItemChildren(list);
            case MarkdownNativeTableBlock table:
                return EnumerateTableChildren(table);
            default:
                return Array.Empty<MarkdownNativeBlock>();
        }
    }

    private static IReadOnlyList<MarkdownNativeBlock> EnumerateTableChildren(MarkdownNativeTableBlock table) {
        if (table == null) {
            return Array.Empty<MarkdownNativeBlock>();
        }

        var children = new List<MarkdownNativeBlock>();
        for (var i = 0; i < table.HeaderCells.Count; i++) {
            children.AddRange(table.HeaderCells[i].Children);
        }

        for (var row = 0; row < table.Rows.Count; row++) {
            for (var column = 0; column < table.Rows[row].Count; column++) {
                children.AddRange(table.Rows[row][column].Children);
            }
        }

        return children;
    }

    private static IEnumerable<MarkdownNativeBlock> EnumerateListItemChildren(MarkdownNativeListBlock list) {
        for (var i = 0; i < list.Items.Count; i++) {
            for (var j = 0; j < list.Items[i].Children.Count; j++) {
                yield return list.Items[i].Children[j];
            }
        }
    }

    private static IEnumerable<MarkdownNativeInline> EnumerateInlines(MarkdownNativeBlock block) {
        foreach (var inline in GetInlineRuns(block)) {
            foreach (var nested in EnumerateInlineAndChildren(inline)) {
                yield return nested;
            }
        }

        foreach (var child in GetChildBlocks(block)) {
            foreach (var inline in EnumerateInlines(child)) {
                yield return inline;
            }
        }
    }

    private static IEnumerable<MarkdownNativeInline> GetInlineRuns(MarkdownNativeBlock block) {
        switch (block) {
            case MarkdownNativeParagraphBlock paragraph:
                return paragraph.InlineRuns;
            case MarkdownNativeHeadingBlock heading:
                return heading.InlineRuns;
            case MarkdownNativeCalloutBlock callout:
                return callout.TitleInlineRuns;
            case MarkdownNativeDetailsBlock details:
                return details.SummaryInlineRuns;
            case MarkdownNativeDefinitionListBlock definitionList:
                return EnumerateDefinitionListTermInlines(definitionList);
            case MarkdownNativeTableBlock table:
                return EnumerateTableInlines(table);
            case MarkdownNativeListBlock list:
                return EnumerateListItemInlines(list);
            default:
                return Array.Empty<MarkdownNativeInline>();
        }
    }

    private static IEnumerable<MarkdownNativeInline> EnumerateTableInlines(MarkdownNativeTableBlock table) {
        for (var i = 0; i < table.HeaderCells.Count; i++) {
            for (var j = 0; j < table.HeaderCells[i].InlineRuns.Count; j++) {
                yield return table.HeaderCells[i].InlineRuns[j];
            }
        }

        for (var row = 0; row < table.Rows.Count; row++) {
            for (var column = 0; column < table.Rows[row].Count; column++) {
                for (var i = 0; i < table.Rows[row][column].InlineRuns.Count; i++) {
                    yield return table.Rows[row][column].InlineRuns[i];
                }
            }
        }
    }

    private static IEnumerable<MarkdownNativeInline> EnumerateDefinitionListTermInlines(MarkdownNativeDefinitionListBlock definitionList) {
        for (var groupIndex = 0; groupIndex < definitionList.Groups.Count; groupIndex++) {
            var group = definitionList.Groups[groupIndex];
            for (var termIndex = 0; termIndex < group.Terms.Count; termIndex++) {
                var term = group.Terms[termIndex];
                for (var inlineIndex = 0; inlineIndex < term.InlineRuns.Count; inlineIndex++) {
                    yield return term.InlineRuns[inlineIndex];
                }
            }
        }
    }

    private static IEnumerable<MarkdownNativeInline> EnumerateListItemInlines(MarkdownNativeListBlock list) {
        for (var i = 0; i < list.Items.Count; i++) {
            for (var j = 0; j < list.Items[i].InlineRuns.Count; j++) {
                yield return list.Items[i].InlineRuns[j];
            }
        }
    }

    private static IEnumerable<MarkdownNativeInline> EnumerateInlineAndChildren(MarkdownNativeInline inline) {
        yield return inline;
        for (var i = 0; i < inline.Children.Count; i++) {
            foreach (var child in EnumerateInlineAndChildren(inline.Children[i])) {
                yield return child;
            }
        }
    }

    private static MarkdownNativeInline? FindInlineAtPosition(MarkdownNativeInline inline, int lineNumber, int columnNumber) {
        if (!inline.ContainsPosition(lineNumber, columnNumber)) {
            return null;
        }

        for (var i = 0; i < inline.Children.Count; i++) {
            var match = FindInlineAtPosition(inline.Children[i], lineNumber, columnNumber);
            if (match != null) {
                return match;
            }
        }

        return inline;
    }

    private static MarkdownNativeInline? FindInlineById(MarkdownNativeInline inline, string id) {
        if (string.Equals(inline.Id, id, StringComparison.Ordinal)) {
            return inline;
        }

        for (var i = 0; i < inline.Children.Count; i++) {
            var match = FindInlineById(inline.Children[i], id);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private bool TryResolveOffsets(MarkdownSourceSpan span, out int startOffset, out int endOffsetInclusive) {
        if (SourceMarkdown.Length == 0) {
            startOffset = 0;
            endOffsetInclusive = -1;
            return false;
        }

        if (span.StartOffset.HasValue && span.EndOffset.HasValue) {
            startOffset = ClampOffset(span.StartOffset.Value);
            endOffsetInclusive = ClampOffset(span.EndOffset.Value);
            return endOffsetInclusive >= startOffset;
        }

        if (span.StartColumn.HasValue && span.EndColumn.HasValue
            && TryGetOffset(span.StartLine, span.StartColumn.Value, out startOffset)
            && TryGetOffset(span.EndLine, span.EndColumn.Value, out endOffsetInclusive)) {
            return endOffsetInclusive >= startOffset;
        }

        if (TryGetLineStartOffset(span.StartLine, out startOffset)
            && TryGetLineEndOffset(span.EndLine, out endOffsetInclusive)) {
            return endOffsetInclusive >= startOffset;
        }

        startOffset = 0;
        endOffsetInclusive = -1;
        return false;
    }

    private int ClampOffset(int offset) {
        if (offset < 0) {
            return 0;
        }

        return offset >= SourceMarkdown.Length ? SourceMarkdown.Length - 1 : offset;
    }

    private bool TryGetOffset(int lineNumber, int columnNumber, out int offset) {
        if (!TryGetLineStartOffset(lineNumber, out var lineStart)) {
            offset = 0;
            return false;
        }

        offset = Math.Min(SourceMarkdown.Length - 1, lineStart + Math.Max(0, columnNumber - 1));
        return true;
    }

    private bool TryGetLineStartOffset(int lineNumber, out int offset) {
        offset = 0;
        if (lineNumber < 1) {
            return false;
        }

        if (lineNumber == 1) {
            return true;
        }

        var currentLine = 1;
        for (var i = 0; i < SourceMarkdown.Length; i++) {
            if (SourceMarkdown[i] != '\n') {
                continue;
            }

            currentLine++;
            if (currentLine == lineNumber) {
                offset = i + 1;
                return offset <= SourceMarkdown.Length;
            }
        }

        return false;
    }

    private bool TryGetLineEndOffset(int lineNumber, out int offset) {
        if (!TryGetLineStartOffset(lineNumber, out var lineStart)) {
            offset = 0;
            return false;
        }

        offset = SourceMarkdown.Length - 1;
        for (var i = lineStart; i < SourceMarkdown.Length; i++) {
            if (SourceMarkdown[i] == '\n') {
                offset = Math.Max(lineStart, i - 1);
                if (offset > lineStart && SourceMarkdown[offset] == '\r') {
                    offset--;
                }

                return true;
            }
        }

        return true;
    }
}
