using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Corpus_Tests {
    [Fact]
    public void Corpus_Samples_Parse_And_Render() {
        // A lightweight "real-world-ish" corpus to catch regressions across lists/tables/quotes/callouts/links.
        // Keep assertions broad to avoid brittle HTML snapshot churn.
        string[] samples = new[] {
            "# Title\n\nParagraph with https://example.com! and `code`.\n",
            "> Quote\n>\n> Continued line\n",
            "> [!NOTE]\n> Callout body line\n>\n> - Item 1\n> - Item 2\n",
            "- Item A\n  - Nested B\n    - Nested C\n\n1. One\n2. Two\n",
            "| A | B |\n|---|---:|\n| x | 1 |\n| y | 2 |\n",
            "Text with [**bold label** link](https://example.com).\n",
            "Footnote ref[^a].\n\n[^a]: Footnote *content*.\n",
            "```csharp\nConsole.WriteLine(\"x\");\n```\n",
            "![alt](/img.png \"t\")\n",
            "Term: Definition\n"
        };

        var readerOptions = new MarkdownReaderOptions { HtmlBlocks = false, InlineHtml = false };
        var htmlOptions = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null };

        for (int i = 0; i < samples.Length; i++) {
            var md = samples[i];
            var doc = MarkdownReader.Parse(md, readerOptions);
            var html = doc.ToHtmlFragment(htmlOptions);

            Assert.False(string.IsNullOrWhiteSpace(html));
            Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
        }
    }
}

