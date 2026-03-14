using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Profile_Tests {
    [Fact]
    public void Portable_Profile_Disables_OfficeImo_Only_Block_Extensions() {
        var options = MarkdownReaderOptions.CreatePortableProfile();

        Assert.False(options.Callouts);
        Assert.False(options.TaskLists);
        Assert.False(options.TocPlaceholders);
        Assert.False(options.Footnotes);
        Assert.False(options.AutolinkUrls);
        Assert.False(options.AutolinkWwwUrls);
        Assert.False(options.AutolinkEmails);
        Assert.True(options.Tables);
        Assert.True(options.DefinitionLists);
        Assert.Empty(options.BlockParserExtensions);
    }

    [Fact]
    public void CommonMark_Profile_Disables_Extensions_Beyond_Core_Syntax() {
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();

        Assert.False(options.FrontMatter);
        Assert.False(options.Callouts);
        Assert.False(options.TaskLists);
        Assert.False(options.Tables);
        Assert.False(options.DefinitionLists);
        Assert.False(options.TocPlaceholders);
        Assert.False(options.Footnotes);
        Assert.False(options.AutolinkUrls);
        Assert.False(options.AutolinkWwwUrls);
        Assert.False(options.AutolinkEmails);
        Assert.True(options.HtmlBlocks);
        Assert.True(options.InlineHtml);
        Assert.Empty(options.BlockParserExtensions);
    }

    [Fact]
    public void Gfm_Profile_Enables_Gfm_Syntax_But_Not_OfficeImo_Extensions() {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();

        Assert.False(options.FrontMatter);
        Assert.False(options.Callouts);
        Assert.True(options.TaskLists);
        Assert.True(options.Tables);
        Assert.False(options.DefinitionLists);
        Assert.False(options.TocPlaceholders);
        Assert.True(options.Footnotes);
        Assert.True(options.AutolinkUrls);
        Assert.True(options.AutolinkWwwUrls);
        Assert.True(options.AutolinkEmails);
        Assert.Single(options.BlockParserExtensions);
        Assert.Equal(MarkdownReaderBuiltInExtensions.FootnotesExtensionName, options.BlockParserExtensions[0].Name);
    }

    [Fact]
    public void CreateProfile_Returns_Expected_Config() {
        var office = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.OfficeIMO);
        var commonMark = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.CommonMark);
        var gfm = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.GitHubFlavoredMarkdown);
        var portable = MarkdownReaderOptions.CreateProfile(MarkdownReaderOptions.MarkdownDialectProfile.Portable);

        Assert.True(office.Callouts);
        Assert.Equal(3, office.BlockParserExtensions.Count);
        Assert.False(commonMark.Callouts);
        Assert.Empty(commonMark.BlockParserExtensions);
        Assert.True(gfm.Tables);
        Assert.Single(gfm.BlockParserExtensions);
        Assert.False(portable.Footnotes);
        Assert.Empty(portable.BlockParserExtensions);
    }

    [Fact]
    public void CommonMark_Profile_Keeps_Toc_And_Footnote_Syntax_As_Literal_Text() {
        const string markdown = """
[TOC]

Lead[^1]

[^1]: Footnote text
""";

        var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());

        Assert.Equal(3, doc.Blocks.Count);

        var tocParagraph = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
        Assert.Equal("\\[TOC\\]", tocParagraph.Inlines.RenderMarkdown());

        var leadParagraph = Assert.IsType<ParagraphBlock>(doc.Blocks[1]);
        Assert.Equal("Lead\\[^1\\]", leadParagraph.Inlines.RenderMarkdown());

        var footnoteParagraph = Assert.IsType<ParagraphBlock>(doc.Blocks[2]);
        Assert.Equal("\\[^1\\]: Footnote text", footnoteParagraph.Inlines.RenderMarkdown());

        Assert.DoesNotContain(doc.Blocks, block => block is FootnoteDefinitionBlock);
    }

    [Fact]
    public void Gfm_Profile_Parses_Tables_TaskLists_And_Footnotes_But_Not_Toc_Placeholders() {
        const string markdown = """
[TOC]

- [x] Done

| Col |
| --- |
| Value |

Lead[^1]

[^1]: Footnote text
""";

        var doc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());

        Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
        Assert.Equal("\\[TOC\\]", Assert.IsType<ParagraphBlock>(doc.Blocks[0]).Inlines.RenderMarkdown());

        var list = Assert.IsType<UnorderedListBlock>(doc.Blocks[1]);
        Assert.True(list.Items[0].IsTask);
        Assert.True(list.Items[0].Checked);

        Assert.IsType<TableBlock>(doc.Blocks[2]);
        Assert.IsType<ParagraphBlock>(doc.Blocks[3]);
        Assert.IsType<FootnoteDefinitionBlock>(doc.Blocks[4]);
    }

    [Fact]
    public void CommonMark_Profile_Can_Opt_Into_Callout_Extension_Explicitly() {
        var options = MarkdownReaderOptions.CreateCommonMarkProfile();
        options.Callouts = true;
        MarkdownReaderBuiltInExtensions.AddCallouts(options);

        var doc = MarkdownReader.Parse("""
> [!NOTE] Example
> Body text
""", options);

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(doc.Blocks));
        Assert.Equal("note", callout.Kind);
        Assert.Equal("Example", callout.TitleInlines.RenderMarkdown());
    }
}
