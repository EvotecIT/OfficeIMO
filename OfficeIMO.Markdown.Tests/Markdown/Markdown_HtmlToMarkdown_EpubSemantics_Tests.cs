using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownHtmlToMarkdownTestsEpubSemantics {
    [Fact]
    public void HtmlToMarkdown_ConvertsEpubFootnotesToTypedMarkdown() {
        const string html = """
<p id="ref-alpha">Text<a epub:type="noteref" role="doc-noteref" href="#note-alpha">1</a>.</p>
<aside epub:type="footnote" role="doc-footnote" id="note-alpha">
  <p>First <strong>note</strong>.</p>
  <pre><code class="language-csharp">Console.WriteLine(1);</code></pre>
  <a epub:type="backlink" role="doc-backlink" href="#ref-alpha">back</a>
</aside>
""";

        MarkdownDoc document = HtmlConversionDocument.Parse(html).ToMarkdownDocument();

        ParagraphBlock paragraph = Assert.IsType<ParagraphBlock>(document.Blocks[0]);
        FootnoteRefInline reference = Assert.Single(paragraph.Inlines.Nodes.OfType<FootnoteRefInline>());
        FootnoteDefinitionBlock definition = Assert.Single(document.Blocks.OfType<FootnoteDefinitionBlock>());

        Assert.Equal("note-alpha", reference.Label);
        Assert.Equal("note-alpha", definition.Label);
        Assert.Contains(definition.ChildBlocks, block => block is ParagraphBlock);
        CodeBlock code = Assert.Single(definition.ChildBlocks.OfType<CodeBlock>());
        Assert.Equal("csharp", code.Language);
        Assert.Contains("Console.WriteLine(1);", code.Content, StringComparison.Ordinal);

        string markdown = document.ToMarkdown();
        Assert.Contains("Text[^note-alpha].", markdown, StringComparison.Ordinal);
        Assert.Contains("[^note-alpha]: First **note**.", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("back", markdown, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlToMarkdown_ConvertsGitHubFootnoteSectionWithoutListScaffolding() {
        const string html = """
<p>Shape<sup id="fnref-shape"><a href="#fn-shape" data-footnote-ref>1</a></sup>.</p>
<section class="footnotes" data-footnotes>
  <ol>
    <li id="fn-shape"><p>Footnote body <a href="#fnref-shape" class="footnote-backref" data-footnote-backref>return</a></p></li>
  </ol>
</section>
""";

        MarkdownDoc document = HtmlConversionDocument.Parse(html).ToMarkdownDocument();

        FootnoteDefinitionBlock definition = Assert.Single(document.Blocks.OfType<FootnoteDefinitionBlock>());
        FootnoteRefInline reference = Assert.Single(
            document.Blocks.OfType<ParagraphBlock>().SelectMany(block => block.Inlines.Nodes).OfType<FootnoteRefInline>());

        Assert.Equal("shape", reference.Label);
        Assert.Equal("shape", definition.Label);
        Assert.DoesNotContain(document.Blocks, block => block is OrderedListBlock);
        Assert.DoesNotContain("return", document.ToMarkdown(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlToMarkdown_KeepsGenericEpubNotesAsOrdinaryContent() {
        const string html = "<aside epub:type=\"note\" id=\"editorial\"><p>Editorial context.</p></aside>";

        MarkdownDoc document = HtmlConversionDocument.Parse(html).ToMarkdownDocument();
        string markdown = document.ToMarkdown();

        Assert.Empty(document.Blocks.OfType<FootnoteDefinitionBlock>());
        Assert.DoesNotContain("[^", markdown, StringComparison.Ordinal);
        Assert.Contains("Editorial context.", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_ProjectsAriaNamesHeadingsListMarkersAndCodeLanguage() {
        const string html = """
<div role="heading" aria-level="4">Accessible heading</div>
<p><a href="note.xhtml" aria-label="Open note"></a></p>
<figure><span id="cover-label" aria-label="Accessible cover chart">ignored</span><img src="images/cover.png" aria-labelledby="cover-label"></figure>
<ol type="A" start="3"><li>Third</li></ol>
<pre data-language="csharp">Console.WriteLine(3);</pre>
""";
        var options = new HtmlToMarkdownOptions { BaseUri = new Uri("https://example.test/book/chapter.xhtml") };

        MarkdownDoc document = HtmlConversionDocument.Parse(html).ToMarkdownDocument(options);

        HeadingBlock heading = Assert.Single(document.Blocks.OfType<HeadingBlock>());
        LinkInline link = Assert.Single(
            document.Blocks.OfType<ParagraphBlock>().SelectMany(block => block.Inlines.Nodes).OfType<LinkInline>());
        ImageBlock image = Assert.Single(document.Blocks.OfType<ImageBlock>());
        OrderedListBlock list = Assert.Single(document.Blocks.OfType<OrderedListBlock>());
        CodeBlock code = Assert.Single(document.Blocks.OfType<CodeBlock>());

        Assert.Equal(4, heading.Level);
        Assert.Equal("Accessible heading", heading.Text);
        Assert.Equal("Open note", link.Text);
        Assert.Equal("https://example.test/book/note.xhtml", link.Url);
        Assert.Equal("Accessible cover chart", image.Alt);
        Assert.Equal("https://example.test/book/images/cover.png", image.Path);
        Assert.Equal(3, list.Start);
        Assert.Equal(MarkdownOrderedListMarkerStyle.UpperAlpha, list.MarkerStyle);
        Assert.Equal("csharp", code.Language);
    }

    [Fact]
    public void HtmlToMarkdown_RejectsFenceSyntaxInCodeLanguageAttributes() {
        const string html = "<pre data-language=\"csharp&#10;```&#10;# injected\">safe()</pre>";

        MarkdownDoc document = HtmlConversionDocument.Parse(html).ToMarkdownDocument();
        CodeBlock code = Assert.Single(document.Blocks.OfType<CodeBlock>());
        string markdown = document.ToMarkdown();

        Assert.Equal(string.Empty, code.Language);
        Assert.Contains("safe()", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("# injected", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesReversedAndValueResetOrderedListMarkers() {
        const string html = """
<ol reversed><li>Three</li><li value="7">Seven</li><li>Six</li></ol>
<ol type="A" start="27"><li>Twenty seven</li><li>Twenty eight</li></ol>
""";

        MarkdownDoc document = HtmlConversionDocument.Parse(html).ToMarkdownDocument();
        OrderedListBlock[] lists = document.Blocks.OfType<OrderedListBlock>().ToArray();

        Assert.Equal(2, lists.Length);
        Assert.True(lists[0].Reversed);
        Assert.Equal(new[] { "3.", "7.", "6." }, lists[0].Items.Select(item => item.MarkerText));
        Assert.Equal(MarkdownOrderedListMarkerStyle.UpperAlpha, lists[1].MarkerStyle);
        Assert.Equal(new[] { "27.", "28." }, lists[1].Items.Select(item => item.MarkerText));
        string markdown = document.ToMarkdown();
        Assert.Contains("3. Three", markdown, StringComparison.Ordinal);
        Assert.Contains("7. Seven", markdown, StringComparison.Ordinal);
        Assert.Contains("6. Six", markdown, StringComparison.Ordinal);
        Assert.Contains("27. Twenty seven", markdown, StringComparison.Ordinal);
        Assert.Contains("28. Twenty eight", markdown, StringComparison.Ordinal);
        Assert.Contains("<ol start=\"3\" reversed>", ((IMarkdownBlock)lists[0]).RenderHtml(), StringComparison.Ordinal);
        MarkdownDoc parsed = MarkdownReader.Parse(markdown);
        OrderedListBlock parsedList = Assert.Single(parsed.Blocks.OfType<OrderedListBlock>());
        Assert.Equal(5, parsedList.Items.Count);
    }
}
