using System;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Html_Encoding_Tests {
    private const string Code = "c.Panel(widget => x < y & z); alias -> target; &amp; &lt; &#62; <br>";
    private const string EncodedCode = "c.Panel(widget =&gt; x &lt; y &amp; z); alias -&gt; target; &amp;amp; &amp;lt; &amp;#62; &lt;br&gt;";

    [Theory]
    [InlineData("`{0}`")]
    [InlineData("- `{0}`")]
    [InlineData("# `{0}`")]
    [InlineData("[`{0}`](https://example.com/?a=1&b=2)")]
    [InlineData("| `{0}` |\n| --- |\n| `{0}` |")]
    [InlineData("| Code |\n| --- |\n| ``a ` {0}`` |")]
    public void CodeSpans_EncodeLiteralTextOnce_InEveryContext(string template) {
        var markdown = string.Format(template, Code);
        var html = Render(markdown);
        var expected = template.Contains("``") ? "a ` " + EncodedCode : EncodedCode;

        Assert.Contains("<code>" + expected + "</code>", html, StringComparison.Ordinal);
        Assert.Equal(template.Contains("| `{0}` |\n") ? 2 : 1, Count(html, "<code>" + expected + "</code>"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ParsedTable_CodeSpans_PreserveLiterals_WithEitherInlineHtmlMode(bool inlineHtml) {
        var html = Render("| `" + Code + "` |\n| --- |\n| `" + Code + "` |", inlineHtml);
        Assert.Equal(2, Count(html, "<code>" + EncodedCode + "</code>"));
    }

    [Fact]
    public void BuilderTable_CodeSpans_KeepLiteralEntitiesAndHtml() {
        var table = new TableBlock();
        table.Headers.Add("`" + Code + "`");
        table.Rows.Add(new[] { "`" + Code + "`" });

        Assert.Equal(2, Count(((IMarkdownBlock)table).RenderHtml(), "<code>" + EncodedCode + "</code>"));
    }

    [Fact]
    public void TableCode_SourceContentAndEscapedPipes_RetainTheirMeaning() {
        const string markdown = "| Code |\n| --- |\n| `x > y &amp; \\| z` |";
        var result = MarkdownReader.ParseWithSyntaxTree(markdown);
        var table = Assert.IsType<TableBlock>(Assert.Single(result.Document.Blocks));
        var code = Assert.IsType<CodeSpanInline>(Assert.Single(table.RowInlines[0][0].Nodes));

        Assert.Equal("x > y &amp; | z", code.Text);
        Assert.Equal(new MarkdownSourceSpan(3, 4, 3, 19), code.ContentSourceSpan);

        var builder = new TableBlock();
        builder.Headers.Add("Code");
        builder.Rows.Add(new[] { @"`x > y &amp; \| z`" });
        Assert.Contains(@"<code>x &gt; y &amp;amp; \| z</code>", ((IMarkdownBlock)builder).RenderHtml(), StringComparison.Ordinal);
    }

    [Fact]
    public void Table_EntitiesAndLinkAttributes_AreEncodedOnce() {
        const string cell = "&amp; &lt; &#62; &amp;amp; [A &amp; B](https://example.com/?a=1&b=2)";
        const string expected = "&amp; &lt; &gt; &amp;amp; <a href=\"https://example.com/?a=1&amp;b=2\">A &amp; B</a>";
        var table = new TableBlock();
        table.Headers.Add("Header");
        table.Rows.Add(new[] { cell });

        Assert.Contains("<td>" + expected + "</td>", Render("| Header |\n| --- |\n| " + cell + " |"), StringComparison.Ordinal);
        Assert.Contains("<td>" + expected + "</td>", ((IMarkdownBlock)table).RenderHtml(), StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RawHtmlAndEscapedMarkup_KeepTheirExistingBoundaries() {
        const string markdown = "| Header |\n| --- |\n| <span data-label=\"A &amp; B\">A &amp; B</span> <b>bold</b><br>next \\<em>literal\\</em> <script>alert(1)</script> |";
        var html = Render(markdown);

        Assert.Contains("<span data-label=\"A &amp; B\">A &amp; B</span>", html, StringComparison.Ordinal);
        Assert.Contains("<b>bold</b><br>next &lt;em&gt;literal&lt;/em&gt;", html, StringComparison.Ordinal);
        Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt;", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("[x](https://example.com \"`\")")]
    [InlineData("![x](https://example.com/a.png \"`\")")]
    [InlineData("[x](https://example.com/`path)")]
    [InlineData("[<u><script>alert(2)</script></u>](https://example.com)")]
    [InlineData("![x](https://example.com/`path.png)")]
    public void Table_BackticksInLinkOrImageMetadata_DoNotEnableRawHtml(string prefix) {
        var cell = prefix + " <script>alert(1)</script> `";
        var builder = new TableBlock();
        builder.Headers.Add("Header");
        builder.Rows.Add(new[] { cell });
        var outputs = new[] { Render("| Header |\n| --- |\n| " + cell + " |"), ((IMarkdownBlock)builder).RenderHtml() };

        foreach (var html in outputs) {
            Assert.DoesNotContain("<script>", html, StringComparison.Ordinal);
            Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt;", html, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void BuilderTable_HtmlFormatting_StillUsesItsRestrictedTagSet() {
        var table = new TableBlock();
        table.Headers.Add("Header");
        table.Rows.Add(new[] { "<span>plain</span> <u>under <script>alert(1)</script></u> <q>quote</q>" });

        var html = ((IMarkdownBlock)table).RenderHtml();
        Assert.Contains("&lt;span&gt;plain&lt;/span&gt;", html, StringComparison.Ordinal);
        Assert.Contains("<u>under &lt;script&gt;alert(1)&lt;/script&gt;</u>", html, StringComparison.Ordinal);
        Assert.Contains("&lt;q&gt;quote&lt;/q&gt;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ParsedTable_RawHtmlAttributes_PreserveLiteralBreakText() {
        const string markdown = "| Header |\n| --- |\n| before <span data-template=\"<br>\">value</span><br>after |";
        var html = Render(markdown);

        Assert.Contains("<span data-template=\"<br>\">value</span>", html, StringComparison.Ordinal);
        Assert.Matches("</span><br\\s*/?>after", html);
    }

    [Theory]
    [InlineData("")]
    [InlineData("- ")]
    public void TableCode_RepeatedBreakText_DoesNotCreateCellBlocks(string prefix) {
        var cell = prefix + "`one<br><br>two`";
        var table = new TableBlock();
        table.Headers.Add("Header");
        table.Rows.Add(new[] { cell });
        var outputs = new[] { Render("| Header |\n| --- |\n| " + cell + " |"), ((IMarkdownBlock)table).RenderHtml() };
        foreach (var html in outputs) {
            Assert.Contains("<code>one&lt;br&gt;&lt;br&gt;two</code>", html, StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("```")]
    [InlineData("~~~")]
    public void Table_FencedCodeContent_Backticks_DoNotSuppressEncodedLineBreaks(string fence) {
        var cell = fence + "text<br>`a<br>b`<br>" + fence;
        var parsed = Assert.IsType<TableBlock>(Assert.Single(MarkdownReader.Parse("| Header |\n| --- |\n| " + cell + " |").Blocks));
        var builder = new TableBlock();
        builder.Headers.Add("Header");
        builder.Rows.Add(new[] { cell });

        foreach (var table in new[] { parsed, builder }) {
            var code = Assert.IsType<CodeBlock>(Assert.Single(table.RowCells[0][0].ChildBlocks));
            Assert.Equal("`a\nb`", code.Content);
        }
    }

    [Fact]
    public void RawHtmlOutsideTables_IsPassedThroughWithoutReencoding() {
        const string raw = "<div data-label=\"A &amp; B\">A &amp; B</div>";
        Assert.Contains(raw, Render(raw), StringComparison.Ordinal);
    }

    private static string Render(string markdown, bool inlineHtml = true) =>
        MarkdownReader.Parse(markdown, new MarkdownLoadOptions { InlineHtml = inlineHtml })
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None });

    private static int Count(string value, string part) => (value.Length - value.Replace(part, string.Empty).Length) / part.Length;
}
