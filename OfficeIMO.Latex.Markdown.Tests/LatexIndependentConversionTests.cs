using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexIndependentConversionTests {
    [Theory]
    [InlineData("small2e.tex")]
    [InlineData("sample2e.tex")]
    public void LaTeXProject_samples_keep_supported_text_and_structured_children(string file) {
        LatexDocument document = LatexDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "latex2e", file));
        LatexToMarkdownResult result = document.ToMarkdownDocumentResult();
        string pdf = PdfReadDocument.Open(document.ToPdfDocumentResult().Value.ToBytes()).ExtractText();
        if (file == "small2e.tex") {
            Assert.Equal(new[] { "Simple Text", "A Warning or Two" }, result.Value.Blocks.OfType<HeadingBlock>().Select(static heading => heading.Text));
            Assert.Single(result.Value.Descendants().OfType<BoldSequenceInline>());
            Assert.Single(result.Value.Descendants().OfType<ItalicSequenceInline>());
            Assert.Contains("this is emphasized", pdf, StringComparison.Ordinal);
            Assert.Contains("this is bold", pdf, StringComparison.Ordinal);
        } else {
            UnorderedListBlock outer = Assert.Single(result.Value.Blocks.OfType<UnorderedListBlock>());
            Assert.Equal(3, outer.Items.Count);
            OrderedListBlock inner = Assert.Single(outer.Items[1].NestedBlocks.OfType<OrderedListBlock>());
            Assert.Equal(2, inner.Items.Count);
            Assert.IsType<ParagraphBlock>(outer.Items[1].NestedBlocks.Last());
            QuoteBlock[] quotes = result.Value.Blocks.OfType<QuoteBlock>().ToArray();
            Assert.Equal(new[] { 1, 2 }, quotes.Select(static quote => quote.ChildBlocks.Count));
            FootnoteDefinitionBlock note = Assert.Single(result.Value.Blocks.OfType<FootnoteDefinitionBlock>());
            Assert.Equal("This is an example of a footnote.", Assert.Single(note.ChildBlocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<MarkdownTextRun>().Single().Text);
            foreach (string text in new[] { "This is the third item", "This is the second paragraph", "This is an example of a footnote" })
                Assert.Contains(text, pdf, StringComparison.Ordinal);
            Assert.Contains(result.Report.Diagnostics, static diagnostic => diagnostic.Feature == "environment:verse");
            Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        }
    }
}
