using System;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_DefinitionList_Html_Tests {
        [Theory]
        [InlineData("Term: Definition with *emphasis* inside.", "<em>emphasis</em>")]
        [InlineData("Term: Includes a [link](https://example.com).", "<a href=\"https://example.com\">link</a>")]
        public void DefinitionList_Renders_Inline_Markup(string markdown, string expectedFragment) {
            var doc = MarkdownReader.Parse(markdown);
            var html = doc.ToHtml();

            Assert.Contains("<dl>", html);

            int ddStart = html.IndexOf("<dd>", StringComparison.OrdinalIgnoreCase);
            Assert.True(ddStart >= 0, "Definition element not found in HTML output.");

            int ddEnd = html.IndexOf("</dd>", ddStart, StringComparison.OrdinalIgnoreCase);
            Assert.True(ddEnd > ddStart, "Closing definition tag not found in HTML output.");

            var inner = html.Substring(ddStart + 4, ddEnd - (ddStart + 4));
            Assert.Contains(expectedFragment, inner);
        }
    }
}
