using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

internal static class MarkdownInvariantAssert {
    public static void SemanticTreeIsWellFormed(MarkdownDoc document) {
        Assert.NotNull(document);
        Assert.Null(document.Parent);
        Assert.Null(document.IndexInParent);
        Assert.Null(document.PreviousSibling);
        Assert.Null(document.NextSibling);
        Assert.Same(document, document.Document);
        Assert.Same(document, document.Root);

        AssertSemanticNode(document, expectedParent: null, expectedDocument: document, expectedRoot: document);
    }

    public static void SyntaxTreeIsWellFormed(MarkdownSyntaxNode root) {
        Assert.NotNull(root);
        Assert.Null(root.Parent);
        Assert.Equal(-1, root.IndexInParent);
        Assert.Null(root.PreviousSibling);
        Assert.Null(root.NextSibling);
        Assert.Same(root, root.Root);

        AssertSyntaxNode(root, expectedParent: null, expectedRoot: root);
    }

    public static void MappedAssociatedObjectsAreConsistent(MarkdownParseResult result) {
        var mappedNodes = result.FinalSyntaxTree
            .DescendantsAndSelf()
            .Where(node => node.AssociatedObject is MarkdownObject markdownObject && ReferenceEquals(markdownObject.Document, result.Document))
            .ToList();

        Assert.NotEmpty(mappedNodes);

        foreach (var syntaxNode in mappedNodes) {
            var markdownObject = Assert.IsAssignableFrom<MarkdownObject>(syntaxNode.AssociatedObject);
            Assert.Same(result.Document, markdownObject.Document);
            Assert.Equal(syntaxNode.SourceSpan, markdownObject.SourceSpan);

            if (!syntaxNode.SourceSpan.HasValue) {
                continue;
            }

            var path = result.FindFinalNodePathContainingSpan(syntaxNode.SourceSpan.Value);
            Assert.NotEmpty(path);
            Assert.Contains(syntaxNode, path);

            var overlapping = result.FindDeepestFinalNodeOverlappingSpan(syntaxNode.SourceSpan.Value);
            Assert.NotNull(overlapping);
            Assert.True(overlapping!.SourceSpan.HasValue);
            Assert.True(overlapping.SourceSpan!.Value.Overlaps(syntaxNode.SourceSpan.Value));
        }
    }

    private static void AssertSemanticNode(
        MarkdownObject node,
        MarkdownObject? expectedParent,
        MarkdownDoc expectedDocument,
        MarkdownObject expectedRoot) {
        Assert.Same(expectedParent, node.Parent);
        Assert.Same(expectedDocument, node.Document);
        Assert.Same(expectedRoot, node.Root);

        var children = node.ChildObjects;
        for (int i = 0; i < children.Count; i++) {
            var child = children[i];

            Assert.Same(node, child.Parent);
            Assert.Equal(i, child.IndexInParent);
            AssertSiblingLinks(children, i, child);

            if (node.SourceSpan.HasValue && child.SourceSpan.HasValue) {
                Assert.True(
                    node.SourceSpan.Value.Contains(child.SourceSpan.Value),
                    $"Expected parent semantic span {node.SourceSpan.Value} to contain child span {child.SourceSpan.Value} for {child.GetType().Name}.");
            }

            var ancestorChain = child.AncestorsAndSelf().ToArray();
            Assert.NotEmpty(ancestorChain);
            Assert.Same(child, ancestorChain[0]);
            Assert.Same(node, ancestorChain[1]);
            Assert.Same(expectedRoot, ancestorChain[^1]);

            AssertSemanticNode(child, node, expectedDocument, expectedRoot);
        }
    }

    private static void AssertSyntaxNode(
        MarkdownSyntaxNode node,
        MarkdownSyntaxNode? expectedParent,
        MarkdownSyntaxNode expectedRoot) {
        Assert.Same(expectedParent, node.Parent);
        Assert.Same(expectedRoot, node.Root);

        for (int i = 0; i < node.Children.Count; i++) {
            var child = node.Children[i];

            Assert.Same(node, child.Parent);
            Assert.Equal(i, child.IndexInParent);
            AssertSiblingLinks(node.Children, i, child);

            if (node.SourceSpan.HasValue && child.SourceSpan.HasValue) {
                Assert.True(
                    node.SourceSpan.Value.Contains(child.SourceSpan.Value),
                    $"Expected parent syntax span {node.SourceSpan.Value} to contain child span {child.SourceSpan.Value} for {child.Kind}.");
            }

            var ancestorChain = child.AncestorsAndSelf().ToArray();
            Assert.NotEmpty(ancestorChain);
            Assert.Same(child, ancestorChain[0]);
            Assert.Same(node, ancestorChain[1]);
            Assert.Same(expectedRoot, ancestorChain[^1]);

            AssertSyntaxNode(child, node, expectedRoot);
        }
    }

    private static void AssertSiblingLinks<TNode>(IReadOnlyList<TNode> siblings, int index, TNode current)
        where TNode : class {
        if (index == 0) {
            Assert.Null(GetPreviousSibling(current));
        } else {
            Assert.Same(siblings[index - 1], GetPreviousSibling(current));
        }

        if (index == siblings.Count - 1) {
            Assert.Null(GetNextSibling(current));
        } else {
            Assert.Same(siblings[index + 1], GetNextSibling(current));
        }
    }

    private static object? GetPreviousSibling<TNode>(TNode node)
        where TNode : class =>
        node switch {
            MarkdownObject markdownObject => markdownObject.PreviousSibling,
            MarkdownSyntaxNode syntaxNode => syntaxNode.PreviousSibling,
            _ => null
        };

    private static object? GetNextSibling<TNode>(TNode node)
        where TNode : class =>
        node switch {
            MarkdownObject markdownObject => markdownObject.NextSibling,
            MarkdownSyntaxNode syntaxNode => syntaxNode.NextSibling,
            _ => null
        };
}
