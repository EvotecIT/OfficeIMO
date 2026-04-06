using System.Collections.Generic;
using System.Linq;
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

    [Theory]
    [MemberData(nameof(RepresentativeMarkdownDocuments))]
    public void ParseWithSyntaxTree_RepresentativeDocuments_Satisfy_TreeInvariants(string markdown) {
        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Theory]
    [MemberData(nameof(RepresentativeTransformedMarkdownDocuments))]
    public void ParseWithSyntaxTreeAndDiagnostics_RepresentativeNestedTransforms_Satisfy_FinalTreeInvariants(string markdown) {
        var options = new MarkdownReaderOptions();
        options.DocumentTransforms.Add(new MergeFirstTwoParagraphsInNestedBlockListsTransform("merged"));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.SyntaxTree);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void FluentDocument_RepresentativeStructure_Satisfies_SemanticTreeInvariants() {
        var document = MarkdownDoc.Create()
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

    private sealed class MergeFirstTwoParagraphsInNestedBlockListsTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
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
