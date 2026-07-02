using System.IO;
using System.Linq;
using System.Text;
// Intentionally avoid heavy regex use; simple scanning is used for resilience and speed.

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren) ParseNestedMarkdownBlocks(
        string markdown,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        int lineOffset) {

        var nestedOptions = CloneOptionsWithoutFrontMatter(options);
        var nestedState = CloneState(state);
        var syntaxChildren = new List<MarkdownSyntaxNode>();
        var nestedDoc = ParseInternal(markdown, nestedOptions, nestedState, allowFrontMatter: false, out _, out _, syntaxChildren, lineOffset: lineOffset, applyDocumentTransforms: false);
        return (nestedDoc.Blocks, syntaxChildren);
    }

    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren) ParseNestedMarkdownBlocks(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        IReadOnlyCollection<int>? suppressedParagraphGenericAttributeStartLines = null) {
        if (sourceLines == null || sourceLines.Count == 0) {
            return (Array.Empty<IMarkdownBlock>(), Array.Empty<MarkdownSyntaxNode>());
        }

        var markdown = string.Join("\n", sourceLines.Select(line => line.Text ?? string.Empty));
        var nestedOptions = CloneOptionsWithoutFrontMatter(options);
        var nestedState = CloneState(state);
        nestedState.SourceLineAbsoluteNumbers = sourceLines.Select(line => line.AbsoluteLine).ToArray();
        nestedState.LazyQuoteContinuationLines.Clear();
        nestedState.QuoteContainerLines.Clear();
        nestedState.SuppressedSetextHeadingUnderlineLines.Clear();
        nestedState.SuppressedParagraphGenericAttributeStartLines.Clear();
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            if (sourceLines[lineIndex].IsLazyQuoteContinuation) {
                nestedState.LazyQuoteContinuationLines.Add(lineIndex);
            }

            if (sourceLines[lineIndex].IsQuoteContainerLine) {
                nestedState.QuoteContainerLines.Add(lineIndex);
            }

            if (ShouldSuppressNestedLazySetextHeadingUnderline(sourceLines, lineIndex)) {
                nestedState.SuppressedSetextHeadingUnderlineLines.Add(lineIndex);
            }
        }

        if (suppressedParagraphGenericAttributeStartLines != null) {
            foreach (var lineIndex in suppressedParagraphGenericAttributeStartLines) {
                nestedState.SuppressedParagraphGenericAttributeStartLines.Add(lineIndex);
            }
        }

        var syntaxChildren = new List<MarkdownSyntaxNode>();
        var nestedDoc = ParseInternal(markdown, nestedOptions, nestedState, allowFrontMatter: false, out _, out _, syntaxChildren, lineOffset: 0, applyDocumentTransforms: false);
        var remappedSyntaxChildren = RemapNestedSyntaxNodes(sourceLines, syntaxChildren);
        var remappedSyntaxTree = BuildDocumentSyntaxTree(remappedSyntaxChildren, nestedDoc);
        SynchronizeOwnedSyntaxCaches(remappedSyntaxTree);
        MarkdownObjectTreeBinder.BindDocument(nestedDoc, remappedSyntaxTree);
        return (nestedDoc.Blocks, remappedSyntaxChildren);
    }

    private static bool ShouldSuppressNestedLazySetextHeadingUnderline(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        int lineIndex) {
        if (sourceLines == null ||
            lineIndex <= 0 ||
            lineIndex >= sourceLines.Count ||
            !sourceLines[lineIndex].IsLazyQuoteContinuation ||
            !TryGetSetextHeadingUnderlineLevel(sourceLines[lineIndex].Text, out int level) ||
            level != 1) {
            return false;
        }

        for (int index = lineIndex - 1; index >= 0; index--) {
            var previous = sourceLines[index].Text;
            if (string.IsNullOrWhiteSpace(previous)) {
                return false;
            }

            if (sourceLines[lineIndex].IsQuoteContainerLine &&
                sourceLines[index].IsQuoteContainerLine) {
                return true;
            }

            var trimmed = previous.TrimStart();
            if (trimmed.StartsWith(">", StringComparison.Ordinal) ||
                TryGetUnorderedListMarkerInfo(previous, out _, out _, out _) ||
                TryGetOrderedListMarkerInfo(previous, out _, out _, out _, out _)) {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> RemapNestedSyntaxNodes(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        IReadOnlyList<MarkdownSyntaxNode> syntaxChildren) {
        if (sourceLines == null || sourceLines.Count == 0 || syntaxChildren == null || syntaxChildren.Count == 0) {
            return syntaxChildren ?? Array.Empty<MarkdownSyntaxNode>();
        }

        var remapped = new List<MarkdownSyntaxNode>(syntaxChildren.Count);
        for (int i = 0; i < syntaxChildren.Count; i++) {
            remapped.Add(RemapNestedSyntaxNode(sourceLines, syntaxChildren[i]));
        }

        return remapped;
    }

    private static MarkdownSyntaxNode RemapNestedSyntaxNode(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownSyntaxNode node) {
        var span = RemapNestedSourceSpan(sourceLines, node.SourceSpan);
        if (node.AssociatedObject is MarkdownObject markdownObject) {
            var attributeSourceText = MarkdownGenericAttributeSourceSpans.GetSourceText(markdownObject);
            var attributeSourceSpan = MarkdownGenericAttributeSourceSpans.GetSourceSpan(markdownObject);
            var remappedAttributeSourceSpan = IsSourceSpanAlreadyMappedToSourceLines(sourceLines, attributeSourceSpan)
                ? attributeSourceSpan
                : RemapNestedSourceSpan(sourceLines, attributeSourceSpan);
            if (!string.IsNullOrEmpty(attributeSourceText) && remappedAttributeSourceSpan.HasValue) {
                MarkdownGenericAttributeSourceSpans.Set(markdownObject, attributeSourceText, remappedAttributeSourceSpan);
            }
        }

        if (node.AssociatedObject is QuoteBlock quoteBlock) {
            quoteBlock.ReplaceMarkerSourceSpans(quoteBlock.MarkerSourceSpans
                .Select(marker => RemapNestedSourceSpan(sourceLines, marker) ?? marker)
                .ToArray());
        }

        if (node.AssociatedObject is ListItem listItem) {
            if (listItem.MarkerSourceSpan.HasValue) {
                listItem.MarkerSourceSpan = RemapNestedSourceSpan(sourceLines, listItem.MarkerSourceSpan) ?? listItem.MarkerSourceSpan;
            }

            if (listItem.TaskMarkerSourceSpan.HasValue) {
                listItem.TaskMarkerSourceSpan = RemapNestedSourceSpan(sourceLines, listItem.TaskMarkerSourceSpan) ?? listItem.TaskMarkerSourceSpan;
            }
        }

        if (node.AssociatedObject is TableBlock tableBlock) {
            RemapTableTokenSourceSpans(sourceLines, tableBlock);
        }

        if (node.AssociatedObject is DefinitionListDefinition definition) {
            RemapDefinitionListSidecarSourceSpans(sourceLines, definition);
        }

        if (node.AssociatedObject is HorizontalRuleBlock horizontalRuleBlock &&
            horizontalRuleBlock.MarkerSourceSpan.HasValue) {
            horizontalRuleBlock.MarkerSourceSpan = RemapNestedSourceSpan(sourceLines, horizontalRuleBlock.MarkerSourceSpan)
                ?? horizontalRuleBlock.MarkerSourceSpan;
        }

        IReadOnlyList<MarkdownSyntaxNode> children = node.Children;
        if (node.Children.Count > 0) {
            var remappedChildren = new List<MarkdownSyntaxNode>(node.Children.Count);
            for (int i = 0; i < node.Children.Count; i++) {
                remappedChildren.Add(RemapNestedSyntaxNode(sourceLines, node.Children[i]));
            }

            children = remappedChildren;
        }

        return new MarkdownSyntaxNode(
            node.Kind,
            span,
            node.Literal,
            children,
            node.AssociatedObject,
            node.CustomKind,
            node.Attributes,
            node.IsGenerated);
    }

    private static void SynchronizeOwnedSyntaxCaches(MarkdownSyntaxNode node) {
        if (node == null) {
            throw new ArgumentNullException(nameof(node));
        }

        switch (node.AssociatedObject) {
            case DefinitionListBlock definitionList:
                definitionList.SyntaxItems.Clear();
                for (int i = 0; i < node.Children.Count; i++) {
                    definitionList.SyntaxItems.Add(node.Children[i]);
                }
                break;

            case ListItem listItem:
                SynchronizeListItemSyntaxChildren(listItem, node.Children);
                break;

            case CodeBlock codeBlock:
                codeBlock.SetFenceTokenSourceSpans(
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeFenceOpening),
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeFenceInfo),
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeContent),
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeFenceClosing));
                break;

            case SemanticFencedBlock semanticFencedBlock:
                semanticFencedBlock.SetFenceTokenSourceSpans(
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeFenceOpening),
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeFenceInfo),
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeContent),
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeFenceClosing));
                break;

            case QuoteBlock quoteBlock:
                quoteBlock.SyntaxChildren = node.Children.Count > 0 ? node.Children : null;
                break;

            case DetailsBlock detailsBlock:
                detailsBlock.SyntaxChildren = GetDetailsBodySyntaxChildren(detailsBlock, node);
                break;

            case CustomContainerBlock customContainerBlock:
                customContainerBlock.SyntaxChildren = GetCustomContainerBodySyntaxChildren(node);
                customContainerBlock.OpeningFenceSourceSpan = GetChildSourceSpan(node, MarkdownSyntaxKind.CustomContainerOpeningFence);
                customContainerBlock.InfoSourceSpan = GetChildSourceSpan(node, MarkdownSyntaxKind.CustomContainerInfo);
                customContainerBlock.ClosingFenceSourceSpan = GetChildSourceSpan(node, MarkdownSyntaxKind.CustomContainerClosingFence);
                break;

            case TableCell tableCell:
                tableCell.SyntaxChildren = node.Children.Count > 0 ? node.Children : null;
                break;
        }

        for (int i = 0; i < node.Children.Count; i++) {
            SynchronizeOwnedSyntaxCaches(node.Children[i]);
        }
    }

    private static MarkdownSourceSpan? GetChildSourceSpan(MarkdownSyntaxNode node, MarkdownSyntaxKind kind) {
        for (int i = 0; i < node.Children.Count; i++) {
            if (node.Children[i].Kind == kind) {
                return node.Children[i].SourceSpan;
            }
        }

        return null;
    }

    private static void SynchronizeListItemSyntaxChildren(ListItem listItem, IReadOnlyList<MarkdownSyntaxNode> syntaxChildren) {
        listItem.SyntaxChildren.Clear();

        for (int i = 0; i < syntaxChildren.Count; i++) {
            listItem.SyntaxChildren.Add(syntaxChildren[i]);
        }
    }

    private static void RemapTableTokenSourceSpans(IReadOnlyList<MarkdownSourceLineSlice> sourceLines, TableBlock tableBlock) {
        if (tableBlock.AlignmentCellSources.Count > 0) {
            tableBlock.SetAlignmentCellSources(tableBlock.AlignmentCellSources
                .Select(source => new TableAlignmentCellSource(
                    source.Markdown,
                    RemapTableSidecarSourceSpan(sourceLines, source.SourceSpan)))
                .ToArray());
        }

        if (tableBlock.PipeSources.Count > 0) {
            tableBlock.SetPipeSources(tableBlock.PipeSources
                .Select(source => new TablePipeSource(
                    source.RowIndex,
                    source.ColumnIndex,
                    RemapTableSidecarSourceSpan(sourceLines, source.SourceSpan)))
                .ToArray());
        }
    }

    private static MarkdownSourceSpan RemapTableSidecarSourceSpan(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownSourceSpan sourceSpan) =>
        RemapNestedSourceSpan(sourceLines, sourceSpan) ?? sourceSpan;

    private static void RemapDefinitionListSidecarSourceSpans(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        DefinitionListDefinition definition) {
        if (definition.BlankLineSourceSpans.Count > 0) {
            definition.ReplaceBlankLineSourceSpans(definition.BlankLineSourceSpans
                .Select(span => RemapNestedSidecarSourceSpan(sourceLines, span))
                .ToArray());
        }

        if (definition.ContinuationIndentSourceSpans.Count > 0) {
            definition.ReplaceContinuationIndentSourceSpans(definition.ContinuationIndentSourceSpans
                .Select(span => RemapNestedSidecarSourceSpan(sourceLines, span))
                .ToArray());
        }
    }

    private static MarkdownSourceSpan RemapNestedSidecarSourceSpan(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownSourceSpan sourceSpan) =>
        IsSourceSpanAlreadyMappedToSourceLines(sourceLines, sourceSpan)
            ? sourceSpan
            : RemapNestedSourceSpan(sourceLines, sourceSpan) ?? sourceSpan;

    private static bool ShouldParseBlockGenericAttributes(MarkdownReaderOptions options, MarkdownReaderState? state) =>
        options?.GenericAttributes == true && state?.SuppressBlockGenericAttributes != true;

    private static bool ShouldParseParagraphGenericAttributes(MarkdownReaderOptions options, MarkdownReaderState? state, int startLineIndex) =>
        ShouldParseBlockGenericAttributes(options, state)
        && (state == null || !state.SuppressedParagraphGenericAttributeStartLines.Contains(startLineIndex));

    private static bool ShouldParseNestedStandaloneGenericAttributes(MarkdownReaderOptions options, MarkdownReaderState? state, int lineIndex) {
        if (options?.GenericAttributes != true) {
            return false;
        }

        if (state?.SuppressBlockGenericAttributes != true) {
            return true;
        }

        return state.QuoteContainerLines.Contains(lineIndex);
    }

    private static bool ShouldParseHeadingGenericAttributes(MarkdownReaderOptions options, MarkdownReaderState? state) =>
        ShouldParseBlockGenericAttributes(options, state) && state?.SuppressHeadingGenericAttributes != true;

    private static bool ShouldSuppressAutoIdentifierForLiteralHeadingGenericAttribute(
        string text,
        MarkdownReaderOptions options,
        MarkdownReaderState? state) =>
        options?.GenericAttributes == true
        && (state?.SuppressHeadingGenericAttributes == true
            || !MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(
            text,
            out _,
            out _,
            out _,
            out _,
            requireLeadingWhitespace: true))
        && MarkdownGenericAttributeParser.HasTrailingAttributeBlockSyntax(text, requireLeadingWhitespace: true);

    private static SuppressHeadingGenericAttributesScope SuppressHeadingGenericAttributesInListItems(MarkdownReaderState state) =>
        new SuppressHeadingGenericAttributesScope(state);

    private readonly struct SuppressHeadingGenericAttributesScope : System.IDisposable {
        private readonly MarkdownReaderState _state;
        private readonly bool _previousValue;

        internal SuppressHeadingGenericAttributesScope(MarkdownReaderState state) {
            _state = state;
            _previousValue = state.SuppressHeadingGenericAttributes;
            state.SuppressHeadingGenericAttributes = true;
        }

        public void Dispose() {
            _state.SuppressHeadingGenericAttributes = _previousValue;
        }
    }

    private static IReadOnlyList<MarkdownSyntaxNode>? GetDetailsBodySyntaxChildren(DetailsBlock detailsBlock, MarkdownSyntaxNode node) {
        if (node.Children.Count == 0) {
            return null;
        }

        var bodyChildren = new List<MarkdownSyntaxNode>();
        for (int i = 0; i < node.Children.Count; i++) {
            var child = node.Children[i];
            if (child.Kind == MarkdownSyntaxKind.DetailsOpeningTag ||
                child.Kind == MarkdownSyntaxKind.DetailsClosingTag ||
                child.AssociatedObject is SummaryBlock) {
                continue;
            }

            bodyChildren.Add(child);
        }

        if (bodyChildren.Count == 0) {
            return null;
        }

        return bodyChildren;
    }

    private static IReadOnlyList<MarkdownSyntaxNode>? GetCustomContainerBodySyntaxChildren(MarkdownSyntaxNode node) {
        if (node.Children.Count == 0) {
            return null;
        }

        var bodyChildren = new List<MarkdownSyntaxNode>();
        for (int i = 0; i < node.Children.Count; i++) {
            var child = node.Children[i];
            if (child.Kind == MarkdownSyntaxKind.CustomContainerOpeningFence ||
                child.Kind == MarkdownSyntaxKind.CustomContainerInfo ||
                child.Kind == MarkdownSyntaxKind.CustomContainerClosingFence) {
                continue;
            }

            bodyChildren.Add(child);
        }

        return bodyChildren.Count > 0 ? bodyChildren : null;
    }

    private static MarkdownSourceSpan? RemapNestedSourceSpan(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownSourceSpan? span) {
        if (!span.HasValue) {
            return null;
        }

        var value = span.Value;
        int startIndex = value.StartLine - 1;
        int endIndex = value.EndLine - 1;
        if (startIndex < 0 || startIndex >= sourceLines.Count || endIndex < 0 || endIndex >= sourceLines.Count) {
            return value;
        }

        int startLine = sourceLines[startIndex].AbsoluteLine;
        int endLine = sourceLines[endIndex].AbsoluteLine;
        if (!value.StartColumn.HasValue || !value.EndColumn.HasValue) {
            return new MarkdownSourceSpan(startLine, endLine);
        }

        int startColumn = sourceLines[startIndex].StartColumn + value.StartColumn.Value - 1;
        int endColumn = sourceLines[endIndex].StartColumn + value.EndColumn.Value - 1;
        return new MarkdownSourceSpan(startLine, startColumn, endLine, endColumn);
    }

    private static bool IsSourceSpanAlreadyMappedToSourceLines(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownSourceSpan? span) {
        if (!span.HasValue || sourceLines == null || sourceLines.Count == 0) {
            return false;
        }

        var value = span.Value;
        if (!TryFindSourceLine(sourceLines, value.StartLine, out var startLine) ||
            !TryFindSourceLine(sourceLines, value.EndLine, out var endLine)) {
            return false;
        }

        if (!value.StartColumn.HasValue || !value.EndColumn.HasValue) {
            return true;
        }

        return value.StartColumn.Value >= startLine.StartColumn
            && value.EndColumn.Value >= endLine.StartColumn;
    }

    private static bool TryFindSourceLine(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        int absoluteLine,
        out MarkdownSourceLineSlice sourceLine) {
        for (int i = 0; i < sourceLines.Count; i++) {
            if (sourceLines[i].AbsoluteLine == absoluteLine) {
                sourceLine = sourceLines[i];
                return true;
            }
        }

        sourceLine = default;
        return false;
    }
}
