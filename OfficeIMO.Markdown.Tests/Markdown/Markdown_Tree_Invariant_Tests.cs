using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Tree_Invariant_Tests {
    public static IEnumerable<object[]> RepresentativeMarkdownDocuments() {
        yield return new object[] {
            """
# Title

Lead with **bold** [docs](https://example.com) and `code`.

- first
- second
"""
        };

        yield return new object[] {
            """
<details open>
<summary>Summary</summary>

> [!WARNING] Watch
> Body

1. first
2. second

</details>
"""
        };

        yield return new object[] {
            """
Term: Intro

  - first
  - second

Lead[^1]

[^1]: first line
  continued

  second paragraph

| Name | Value |
| --- | ---: |
| One | 1 |
"""
        };
    }

    public static IEnumerable<object[]> RepresentativeTransformedMarkdownDocuments() {
        yield return new object[] {
            """
> - alpha
>
>   beta
"""
        };

        yield return new object[] {
            """
> [!NOTE] Title
> - alpha
>
>   beta
"""
        };

        yield return new object[] {
            """
<details>
<summary>Summary</summary>

- alpha

  beta
</details>
"""
        };

        yield return new object[] {
            """
Lead[^1]

[^1]:
  - alpha

    beta
"""
        };
    }

    [Fact]
    public void Appending_blocks_one_at_a_time_keeps_every_tree_position_consistent() {
        var parsed = MarkdownDocument.Parse("Moved **paragraph** with [a link](https://example.com).\n\n> quoted\n");
        var document = MarkdownDocument.Create()
            .H1("Title")
            .P(inlines => inlines.Text("Lead ").Bold("bold").Text(" and ").Code("code"))
            .Ul(list => list.Item("first").ItemTask("second", done: true))
            .Code("csharp", "var x = 1;")
            .Quote("quoted line")
            .Ol(list => list.Item("one").Item("two"))
            .Table(table => table.Headers("Name", "Value").Row("One", "1"));
        foreach (IMarkdownBlock block in parsed.Blocks.ToArray()) {
            document.Add(block);
        }

        MarkdownInvariantAssert.SemanticTreeIsWellFormed(document);

        // Front matter becomes the first child, so every block after it moves.
        document.FrontMatter(new { title = "Report" });
        document.P("After front matter");

        MarkdownInvariantAssert.SemanticTreeIsWellFormed(document);
    }

    [Fact]
    public void Editing_the_last_block_before_appending_the_next_keeps_its_new_inlines_bound() {
        var document = MarkdownDocument.Create().P("before");
        var paragraph = (ParagraphBlock)document.Blocks[document.Blocks.Count - 1];
        paragraph.Inlines.Bold("after");

        document.P("next");

        MarkdownInvariantAssert.SemanticTreeIsWellFormed(document);
        var bold = Assert.IsType<BoldInline>(paragraph.Inlines.Nodes[paragraph.Inlines.Nodes.Count - 1]);
        Assert.Same(document, bold.Document);
    }

    [Fact]
    public void Inlines_added_to_an_earlier_block_are_bound_as_they_are_added() {
        var document = MarkdownDocument.Create().P("first").P("second");
        var first = (ParagraphBlock)document.Blocks[0];

        first.Inlines.Italic("later");
        var italic = Assert.IsType<ItalicInline>(first.Inlines.Nodes[first.Inlines.Nodes.Count - 1]);
        Assert.Same(document, italic.Document);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(document);

        first.Inlines.ReplaceItems(new IMarkdownInline[] { new MarkdownTextRun("replaced"), new BoldInline("bold") });
        Assert.Same(document, ((MarkdownObject)first.Inlines.Nodes[1]).Document);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(document);
    }

    [Fact]
    public void Append_after_a_block_whose_children_failed_to_bind_binds_them() {
        var document = MarkdownDocument.Create().P("lead");
        var panel = new ThrowingPanelBlock(new ParagraphBlock(new InlineSequence().Text("inside"))) { ThrowOnChildBlocks = true };

        Assert.Throws<InvalidOperationException>(() => document.Add(panel));
        panel.ThrowOnChildBlocks = false;
        document.P("after");

        MarkdownInvariantAssert.SemanticTreeIsWellFormed(document);
        Assert.Same(panel, ((MarkdownObject)panel.ChildBlocks[0]).Parent);
    }

    private sealed class ThrowingPanelBlock : MarkdownBlock, IMarkdownBlock, IChildMarkdownBlockContainer {
        private readonly IMarkdownBlock[] _children;

        internal ThrowingPanelBlock(params IMarkdownBlock[] children) => _children = children;

        internal bool ThrowOnChildBlocks { get; set; }

        public IReadOnlyList<IMarkdownBlock> ChildBlocks =>
            ThrowOnChildBlocks ? throw new InvalidOperationException("Children are not available yet.") : _children;

        public string RenderMarkdown() => string.Empty;

        public string RenderHtml() => string.Empty;
    }

    [Theory]
    [MemberData(nameof(RepresentativeMarkdownDocuments))]
    public void ParseWithSyntaxTree_RepresentativeDocuments_Satisfy_TreeInvariants(string markdown) {
        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTree(markdown);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Theory]
    [MemberData(nameof(RepresentativeTransformedMarkdownDocuments))]
    public void ParseWithSyntaxTreeAndDiagnostics_RepresentativeNestedTransforms_Satisfy_FinalTreeInvariants(string markdown) {
        var options = new MarkdownLoadOptions();
        options.DocumentTransforms.Add(new MergeFirstTwoParagraphsInNestedBlockListsTransform("merged"));

        var result = OfficeIMO.Markdown.MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.SyntaxTree);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void FluentDocument_RepresentativeStructure_Satisfies_SemanticTreeInvariants() {
        var document = MarkdownDocument.Create()
            .H1("Title")
            .P("Lead paragraph")
            .Ul(list => {
                list.Item("first");
                list.Item("second");
            })
            .Table(table => {
                table.Headers("Name", "Value");
                table.Row("One", "1");
            });

        MarkdownInvariantAssert.SemanticTreeIsWellFormed(document);
    }

    [Fact]
    public void ReaderDocumentTransforms_ObserveCurrentTreeBindings() {
        var probe = new CaptureTreeBindingTransform();
        var options = new MarkdownLoadOptions();
        options.DocumentTransforms.Add(new AppendParagraphTransform());
        options.DocumentTransforms.Add(probe);

        OfficeIMO.Markdown.MarkdownDocument.Parse("First paragraph", options);

        Assert.True(probe.ObservedCurrentBindings);
    }

    private sealed class AppendParagraphTransform : IMarkdownDocumentTransform {
        public MarkdownDocument Transform(MarkdownDocument document, MarkdownDocumentTransformContext context) {
            document.Add(new ParagraphBlock(new InlineSequence().Text("Appended")));
            return document;
        }
    }

    private sealed class CaptureTreeBindingTransform : IMarkdownDocumentTransform {
        public bool ObservedCurrentBindings { get; private set; }

        public MarkdownDocument Transform(MarkdownDocument document, MarkdownDocumentTransformContext context) {
            ObservedCurrentBindings = document.Blocks
                .Select((block, index) => (Block: block as MarkdownObject, Index: index))
                .All(item => item.Block?.Parent == document && item.Block.IndexInParent == item.Index);
            return document;
        }
    }

    private sealed class MergeFirstTwoParagraphsInNestedBlockListsTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDocument Transform(MarkdownDocument document, MarkdownDocumentTransformContext context) {
            MarkdownDocumentBlockListExpander.RewriteDocument(document, context, (blocks, _) => {
                if (blocks.Count >= 2
                    && blocks[0] is ParagraphBlock
                    && blocks[1] is ParagraphBlock) {
                    return new List<IMarkdownBlock> {
                        new ParagraphBlock(new InlineSequence().Text(text))
                    };
                }

                return blocks.ToList();
            });

            return document;
        }
    }
}
