using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexIndependentConversionTests {
    [Theory]
    [InlineData("small2e.tex")]
    [InlineData("sample2e.tex")]
    public void LaTeXProject_samples_keep_supported_text_and_structured_children(string file) {
        LatexDocument document = LatexDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "latex2e", file));
        var result = document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport1);
        var resultReport = (LatexToMarkdownReport)conversionReport1;
        string pdf = PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).ExtractText();
        if (file == "small2e.tex") {
            Assert.Equal(new[] { "Simple Text", "A Warning or Two" }, result.Blocks.OfType<HeadingBlock>().Select(static heading => heading.Text));
            Assert.Single(result.Descendants().OfType<BoldSequenceInline>());
            Assert.Single(result.Descendants().OfType<ItalicSequenceInline>());
            Assert.Contains("this is emphasized", pdf, StringComparison.Ordinal);
            Assert.Contains("this is bold", pdf, StringComparison.Ordinal);
        } else {
            UnorderedListBlock outer = Assert.Single(result.Blocks.OfType<UnorderedListBlock>());
            Assert.Equal(3, outer.Items.Count);
            OrderedListBlock inner = Assert.Single(outer.Items[1].NestedBlocks.OfType<OrderedListBlock>());
            Assert.Equal(2, inner.Items.Count);
            Assert.IsType<ParagraphBlock>(outer.Items[1].NestedBlocks.Last());
            QuoteBlock[] quotes = result.Blocks.OfType<QuoteBlock>().ToArray();
            Assert.Equal(new[] { 1, 2 }, quotes.Select(static quote => quote.ChildBlocks.Count));
            FootnoteDefinitionBlock note = Assert.Single(result.Blocks.OfType<FootnoteDefinitionBlock>());
            Assert.Equal("This is an example of a footnote.", Assert.Single(note.ChildBlocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<MarkdownTextRun>().Single().Text);
            foreach (string text in new[] { "This is the third item", "This is the second paragraph", "This is an example of a footnote" })
                Assert.Contains(text, pdf, StringComparison.Ordinal);
            Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Feature == "environment:verse");
            Assert.ThrowsAny<OfficeConversionException>(() => resultReport.ThrowIfLoss());
        }
    }
}
