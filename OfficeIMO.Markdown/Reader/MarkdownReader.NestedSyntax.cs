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
        MarkdownReaderState state) {
        if (sourceLines == null || sourceLines.Count == 0) {
            return (Array.Empty<IMarkdownBlock>(), Array.Empty<MarkdownSyntaxNode>());
        }

        var markdown = string.Join("\n", sourceLines.Select(line => line.Text ?? string.Empty));
        var nestedOptions = CloneOptionsWithoutFrontMatter(options);
        var nestedState = CloneState(state);
        var syntaxChildren = new List<MarkdownSyntaxNode>();
        var nestedDoc = ParseInternal(markdown, nestedOptions, nestedState, allowFrontMatter: false, out _, out _, syntaxChildren, lineOffset: 0, applyDocumentTransforms: false);
        var remappedSyntaxChildren = RemapNestedSyntaxNodes(sourceLines, syntaxChildren);
        var remappedSyntaxTree = BuildDocumentSyntaxTree(remappedSyntaxChildren, nestedDoc);
        SynchronizeOwnedSyntaxCaches(remappedSyntaxTree);
        MarkdownObjectTreeBinder.BindDocument(nestedDoc, remappedSyntaxTree);
        return (nestedDoc.Blocks, remappedSyntaxChildren);
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

        IReadOnlyList<MarkdownSyntaxNode> children = node.Children;
        if (node.Children.Count > 0) {
            var remappedChildren = new List<MarkdownSyntaxNode>(node.Children.Count);
            for (int i = 0; i < node.Children.Count; i++) {
                remappedChildren.Add(RemapNestedSyntaxNode(sourceLines, node.Children[i]));
            }

            children = remappedChildren;
        }

        return new MarkdownSyntaxNode(node.Kind, span, node.Literal, children, node.AssociatedObject, node.CustomKind);
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
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeFenceClosing));
                break;

            case SemanticFencedBlock semanticFencedBlock:
                semanticFencedBlock.SetFenceTokenSourceSpans(
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeFenceOpening),
                    GetChildSourceSpan(node, MarkdownSyntaxKind.CodeFenceClosing));
                break;

            case QuoteBlock quoteBlock:
                quoteBlock.SyntaxChildren = node.Children.Count > 0 ? node.Children : null;
                break;

            case DetailsBlock detailsBlock:
                detailsBlock.SyntaxChildren = GetDetailsBodySyntaxChildren(detailsBlock, node);
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

        var blockChildrenCount = listItem.BlockChildren.Count;
        var ownedChildCount = Math.Min(blockChildrenCount, syntaxChildren.Count);
        for (int i = 0; i < ownedChildCount; i++) {
            listItem.SyntaxChildren.Add(syntaxChildren[i]);
        }
    }

    private static IReadOnlyList<MarkdownSyntaxNode>? GetDetailsBodySyntaxChildren(DetailsBlock detailsBlock, MarkdownSyntaxNode node) {
        if (node.Children.Count == 0) {
            return null;
        }

        var bodyStartIndex = detailsBlock.Summary != null && node.Children.Count > 0 ? 1 : 0;
        if (bodyStartIndex >= node.Children.Count) {
            return null;
        }

        var bodyChildren = new MarkdownSyntaxNode[node.Children.Count - bodyStartIndex];
        for (int i = bodyStartIndex; i < node.Children.Count; i++) {
            bodyChildren[i - bodyStartIndex] = node.Children[i];
        }

        return bodyChildren;
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
}
