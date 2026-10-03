using OfficeIMO.DocBook;
using OfficeIMO.Markdown;
using OfficeIMO.Reader.DocBook;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderDocBookIntegrityContractsTests {
    [Theory]
    [InlineData("*important*")]
    [InlineData("[shown](https://example.com)")]
    [InlineData("1. literal text")]
    [InlineData("# literal heading")]
    [InlineData("<b>literal HTML</b>")]
    [InlineData("a | b _c_ `d`")]
    public void OrdinaryDocBookTextStaysLiteralAfterMarkdownParsing(string text) {
        var document = DocBookDocument.CreateArticle();
        document.AddParagraph(text);
        var chunk = Assert.Single(DocBookReaderAdapter.Read(document));
        Assert.Equal(text, chunk.Text);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(MarkdownReader.Parse(chunk.Markdown!).Blocks));
        Assert.Equal(text, string.Concat(paragraph.Inlines.Nodes.Cast<MarkdownTextRun>().Select(run => run.Text)));
    }

    [Fact]
    public void EscapedLiteralTextCanBeReassembledAcrossTinyChunkBudgets() {
        const string text = "*[x](y)* 🙂 1. text";
        var document = DocBookDocument.CreateArticle();
        document.AddParagraph(text);
        var chunks = DocBookReaderAdapter.Read(document, readerOptions: new ReaderOptions { MaxChars = 1 }).ToArray();
        Assert.All(chunks, chunk => Assert.True(chunk.Markdown!.Length <= 2)); // A surrogate pair stays intact.
        Assert.Equal(text, string.Concat(chunks.Select(chunk => chunk.Text)));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(MarkdownReader.Parse(string.Concat(chunks.Select(chunk => chunk.Markdown))).Blocks));
        Assert.Equal(text, string.Concat(paragraph.Inlines.Nodes.Cast<MarkdownTextRun>().Select(run => run.Text)));
    }

    [Fact]
    public void NativeInlineEmphasisAndLiteralTextHaveDistinctMarkdownSemantics() {
        const string xml = "<article xmlns='http://docbook.org/ns/docbook' version='5.2'><para>plain <emphasis>italic</emphasis><emphasis role='bold'>bold</emphasis> <literal>*code*</literal></para></article>";
        var chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(xml)).ToArray();
        Assert.Equal("plain italicbold *code*", string.Concat(chunks.Select(chunk => chunk.Text)));
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(MarkdownReader.Parse(string.Concat(chunks.Select(chunk => chunk.Markdown))).Blocks));
        Assert.Contains(paragraph.Inlines.Nodes, node => node is ItalicSequenceInline || node is ItalicInline);
        Assert.Contains(paragraph.Inlines.Nodes, node => node is BoldSequenceInline || node is BoldInline);
        Assert.Contains(paragraph.Inlines.Nodes, node => node is CodeSpanInline code && code.Text == "*code*");
    }

    [Fact]
    public void FlattenedTableCellsRetainParagraphBoundaries() {
        const string xml = "<article xmlns='http://docbook.org/ns/docbook' version='5.2'><informaltable><tgroup cols='1'><tbody><row><entry><para>one</para><para>two</para></entry></row></tbody></tgroup></informaltable></article>";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        var result = DocBookReaderAdapter.ReadDocument(stream);
        Assert.Equal("one\ntwo", Assert.Single(Assert.Single(result.Tables).Rows)[0]);
    }

    [Fact]
    public void EmphasisWithBoundaryWhitespaceUsesInlineHtmlWithoutAddingVisibleDelimiters() {
        const string xml = "<article xmlns='http://docbook.org/ns/docbook' version='5.2'><para><emphasis> italic </emphasis></para></article>";
        var chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(xml)).ToArray();
        Assert.Equal(" italic ", string.Concat(chunks.Select(chunk => chunk.Text)));
        Assert.Equal("<em> italic </em>", string.Concat(chunks.Select(chunk => chunk.Markdown)));
    }

    [Theory]
    [InlineData("before <link xlink:href='https://example.com'>middle</link>", false, 500)]
    [InlineData("<link xlink:href='https://example.com'>middle</link> after", false, 500)]
    [InlineData("before <link xlink:href='https://example.com'>middle</link> after", true, 500)]
    [InlineData("<emphasis role='bold'>before <link xlink:href='https://example.com'>middle</link></emphasis>", false, 500)]
    [InlineData("before <link xlink:href='https://example.com'>middle</link>", false, 1)]
    public void EmphasisAroundLinksPreservesEveryFragmentAndItsVisibleText(string content, bool bold, int maxChars) {
        string xml = "<article xmlns='http://docbook.org/ns/docbook' xmlns:xlink='http://www.w3.org/1999/xlink' version='5.2'><para><emphasis" +
            (bold ? " role='bold'" : "") + ">" + content + "</emphasis></para></article>";
        var chunks = DocBookReaderAdapter.Read(DocBookDocument.Parse(xml), readerOptions: new ReaderOptions { MaxChars = maxChars }).ToArray();
        var markdown = MarkdownReader.Parse(string.Concat(chunks.Select(chunk => chunk.Markdown)));
        string html = markdown.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        var paragraph = System.Xml.Linq.XElement.Parse(html.Trim());

        Assert.Equal(string.Concat(chunks.Select(chunk => chunk.Text)), paragraph.Value);
        Assert.All(paragraph.DescendantNodes().OfType<System.Xml.Linq.XText>(), text =>
            Assert.Contains(text.Ancestors(), element => element.Name.LocalName == (bold ? "strong" : "em")));
        var link = Assert.Single(paragraph.Descendants("a"));
        Assert.Equal("https://example.com", (string?)link.Attribute("href"));
        Assert.Equal("middle", link.Value);
    }
}
