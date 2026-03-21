using System.Collections.Generic;
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

    [Theory]
    [MemberData(nameof(RepresentativeMarkdownDocuments))]
    public void ParseWithSyntaxTree_RepresentativeDocuments_Satisfy_TreeInvariants(string markdown) {
        var result = MarkdownReader.ParseWithSyntaxTree(markdown);

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
}
