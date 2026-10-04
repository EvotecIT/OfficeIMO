namespace OfficeIMO.AsciiDoc.Markdown.Tests;

public sealed class AsciiDocSectionReferenceTests {
    [Fact]
    public void GeneratedCrossReferencesReachHeadingsInHtmlAndPreserveClasses() {
        AsciiDocDocument document = AsciiDocDocument.Parse("[.wide]\n== A *Bold* Title\n\n== A *Bold* Title\n\nSee <<_a_bold_title>> and xref:_a_bold_title_2[].\n");
        AsciiDocToMarkdownResult result = document.ToMarkdownDocumentResult();
        HeadingBlock[] headings = result.Value.Blocks.OfType<HeadingBlock>().ToArray();
        Assert.Equal(new[] { "_a_bold_title", "_a_bold_title_2" }, headings.Select(heading => heading.Attributes.ElementId));
        Assert.Contains("wide", headings[0].Attributes.Classes);
        string html = result.Value.ToHtmlFragment();
        Assert.Contains("id=\"_a_bold_title\"", html);
        Assert.Contains("id=\"_a_bold_title_2\"", html);
        Assert.Contains("href=\"#_a_bold_title\"", html);
        Assert.Contains("href=\"#_a_bold_title_2\"", html);
        Assert.Contains("[A Bold Title](#_a_bold_title)", result.Value.ToMarkdown());
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "ADOCREF003");
    }

    [Fact]
    public void ExplicitInlineHeadingAnchorHasOneHtmlTargetAndDisabledIdsStayDisabled() {
        AsciiDocDocument document = AsciiDocDocument.Parse("== Inline [[inline-id]]\n\n:!sectids:\n\n== Disabled\n\nSee <<inline-id>>.\n");
        AsciiDocToMarkdownResult result = document.ToMarkdownDocumentResult();
        HeadingBlock[] headings = result.Value.Blocks.OfType<HeadingBlock>().ToArray();
        Assert.Equal("inline-id", headings[0].Attributes.ElementId);
        Assert.Null(headings[1].Attributes.ElementId);
        string html = result.Value.ToHtmlFragment();
        Assert.Equal(1, html.Split(new[] { "id=\"inline-id\"" }, StringSplitOptions.None).Length - 1);
        Assert.Contains("href=\"#inline-id\"", html);
        Assert.DoesNotContain("id=\"disabled\"", html);
    }
}
