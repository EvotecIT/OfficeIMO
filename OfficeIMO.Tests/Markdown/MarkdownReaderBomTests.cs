using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.Markdown;

public class MarkdownReaderBomTests {
    [Fact]
    public void Bom_With_Heading_Parses_As_Heading() {
        var md = "\uFEFF# Heading";
        var doc = MarkdownReader.Parse(md);
        Assert.Single(doc.Blocks);
        var h = Assert.IsType<HeadingBlock>(doc.Blocks[0]);
        Assert.Equal(1, h.Level);
        Assert.Equal("Heading", h.Text);
    }

    [Fact]
    public void Bom_With_Html_Parses_As_HtmlRawBlock() {
        var md = "\uFEFF<div>content</div>";
        var doc = MarkdownReader.Parse(md);
        Assert.Single(doc.Blocks);
        var raw = Assert.IsType<HtmlRawBlock>(doc.Blocks[0]);
        Assert.Equal("<div>content</div>", raw.Html.Trim());
    }

    [Fact]
    public void Bom_With_Paragraph_Parses_As_Paragraph() {
        var md = "\uFEFFPlain text";
        var doc = MarkdownReader.Parse(md);
        Assert.Single(doc.Blocks);
        var p = Assert.IsType<ParagraphBlock>(doc.Blocks[0]);
        var items = p.Inlines.Items;
        Assert.Single(items);
        var t = Assert.IsType<TextRun>(items[0]);
        Assert.Equal("Plain text", t.Text);
    }

    [Fact]
    public void No_Bom_Existing_Behavior_Unchanged() {
        var md = "# Heading";
        var doc = MarkdownReader.Parse(md);
        var h = Assert.IsType<HeadingBlock>(doc.Blocks[0]);
        Assert.Equal(1, h.Level);
        Assert.Equal("Heading", h.Text);
    }

    [Fact]
    public void Bom_Only_Produces_No_Blocks() {
        var md = "\uFEFF";
        var doc = MarkdownReader.Parse(md);
        Assert.Empty(doc.Blocks);
    }
}

