using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Refs_Footnotes_Tests {
        [Fact]
        public void Reference_Links_Are_Resolved() {
            var md = string.Join("\n", new[] {
                "See [Docs][docs] and [Site][site].",
                "",
                "[docs]: https://evotec.xyz \"Docs\"",
                "[site]: <https://example.com> \"Site\""
            });
            var doc = MarkdownReader.Parse(md);
            var outMd = doc.ToMarkdown();
            // Either inline links or preserved, accept either; primarily ensure resolution in HTML
            var html = doc.ToHtml();
            Assert.Contains("https://evotec.xyz", html);
            Assert.Contains("https://example.com", html);
            Assert.DoesNotContain("[docs]:", outMd); // definitions consumed
        }

        [Fact]
        public void Reference_Links_With_Nested_Label_Text_Are_Resolved() {
            var md = string.Join("\n", new[] {
                "See [Docs [API]][docs].",
                "",
                "[docs]: https://evotec.xyz"
            });

            var html = MarkdownReader.Parse(md).ToHtml();

            Assert.Contains("href=\"https://evotec.xyz\"", html);
            Assert.Contains(">Docs [API]<", html);
        }

        [Fact]
        public void Footnote_Refs_And_Definitions_RoundTrip() {
            var md = string.Join("\n", new[] {
                "Hello[^1] world.",
                "",
                "[^1]: A note",
            });
            var doc = MarkdownReader.Parse(md);
            var outMd = doc.ToMarkdown();
            Assert.Contains("[^1]", outMd);
            Assert.Contains("[^1]: A note", outMd);
            var html = doc.ToHtml();
            Assert.Contains("id=\"fnref:1\"", html);
            Assert.Contains("id=\"fn:1\"", html);
        }

        [Fact]
        public void Standalone_Footnote_Block_Reuses_Parsed_Paragraphs_When_Available() {
            var md = string.Join("\n", new[] {
                "Lead[^1]",
                "",
                "[^1]: First *line*",
                "",
                "  Second [link](https://example.com)"
            });

            var doc = MarkdownReader.Parse(md);
            var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(doc.Blocks, b => b is FootnoteDefinitionBlock));

            var html = ((IMarkdownBlock)footnote).RenderHtml();

            Assert.Contains("<em>line</em>", html, StringComparison.Ordinal);
            Assert.Contains("href=\"https://example.com\"", html, StringComparison.Ordinal);
            Assert.Equal(2, footnote.Paragraphs.Count);
        }
    }
}
