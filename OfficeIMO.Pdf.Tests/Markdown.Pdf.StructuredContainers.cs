using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class MarkdownPdfTests {
    [Fact]
    public void Callout_with_image_and_following_panel_content_renders_its_title_once() {
        string markdown = "> [!NOTE] Single unique title\n>\n> ![Pixel](" + CreateMinimalRgbPngDataUri() + ")\n>\n> Following content.\n";
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(MarkdownReader.Parse(markdown).ToPdfBytes()).ExtractText();
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(text, "Single unique title").Cast<System.Text.RegularExpressions.Match>());
        Assert.Contains("Following content.", text);
    }

    [Theory]
    [InlineData("4.", "5.", "6.", "")]
    [InlineData("-", "-", "-", "")]
    [InlineData("- [x]", "-", "- [ ]", "")]
    [InlineData("-", "-", "-", "quote")]
    [InlineData("-", "-", "-", "footnote")]
    public void List_children_precede_their_next_outer_sibling(string first, string second, string third, string wrapper) {
        MarkdownDoc body = MarkdownReader.Parse($"{first} Outer first\n\n{second} Outer second\n\n    - Nested alpha\n    - Nested beta\n\n    Continued second.\n\n{third} Outer third\n");
        IMarkdownBlock block = Assert.Single(body.Blocks);
        if (wrapper == "quote") {
            var quote = new QuoteBlock();
            quote.ChildBlocks.AddRange(body.Blocks);
            block = quote;
        }
        if (wrapper == "footnote") block = new FootnoteDefinitionBlock("note", body.Blocks);
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(MarkdownDoc.Create().Add(block).ToPdfBytes()).ExtractText();
        int previous = -1;
        foreach (string marker in new[] { "Outer first", "Outer second", "Nested alpha", "Nested beta", "Continued second.", "Outer third" }) {
            int current = text.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(current > previous, text);
            previous = current;
        }
        if (first == "4.") foreach (string marker in new[] { "4.", "5.", "6." }) Assert.Contains(marker, text);
    }

    [Fact]
    public void Description_definitions_render_owned_blocks_and_shared_terms_without_markdown_literals() {
        var definition = new DefinitionListBlock();
        definition.AddGroup(new DefinitionListGroup(
            new[] { MarkdownReader.ParseInlineText("First term"), MarkdownReader.ParseInlineText("Second term") },
            new[] { new DefinitionListDefinition(MarkdownReader.Parse("First paragraph.\n\nSecond paragraph.\n\n> Quoted content\n\n- Nested item").Blocks) }));
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(MarkdownDoc.Create().Add(definition).ToPdfBytes()).ExtractText();
        foreach (string marker in new[] { "First term", "Second term", "First paragraph.", "Second paragraph.", "Quoted content", "Nested item" })
            Assert.Single(System.Text.RegularExpressions.Regex.Matches(text, System.Text.RegularExpressions.Regex.Escape(marker)).Cast<System.Text.RegularExpressions.Match>());
        Assert.DoesNotContain("> Quoted", text);
        Assert.DoesNotContain("- Nested", text);
    }
}
