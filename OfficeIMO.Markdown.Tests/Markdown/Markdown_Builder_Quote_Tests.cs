using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Builder_Quote_Tests {
        private static string Normalize(string value) => value.Replace("\r\n", "\n");

        [Fact]
        public void QuoteBuilder_Renders_Simple_Lines() {
            var doc = MarkdownDoc.Create().Quote(q => q.Line("Line 1").Line("Line 2"));

            string markdown = Normalize(doc.ToMarkdown());
            Assert.Equal("> Line 1\n> Line 2\n", markdown);

            string html = doc.Blocks[0].RenderHtml();
            Assert.Equal("<blockquote><p>Line 1<br/>Line 2</p></blockquote>", html);

            var singleLine = MarkdownDoc.Create().Quote("Single line");
            string singleMarkdown = Normalize(singleLine.ToMarkdown());
            Assert.Equal("> Single line\n", singleMarkdown);
            Assert.Equal("<blockquote><p>Single line</p></blockquote>", singleLine.Blocks[0].RenderHtml());
        }

        [Fact]
        public void QuoteBuilder_Renders_Nested_Blocks() {
            var doc = MarkdownDoc.Create().Quote(q => q
                .Line("Intro")
                .Quote(inner => inner.Line("Inner line 1").Line("Inner line 2"))
                .P(p => p.Text("Conclusion section"))
                .Line("Closing note"));

            string markdown = Normalize(doc.ToMarkdown());
            Assert.Equal(
                "> Intro\n> \n> > Inner line 1\n> > Inner line 2\n> \n> Conclusion section\n> \n> Closing note\n",
                markdown);

            string html = doc.Blocks[0].RenderHtml();
            Assert.Equal(
                "<blockquote><p>Intro</p><blockquote><p>Inner line 1<br/>Inner line 2</p></blockquote><p>Conclusion section</p><p>Closing note</p></blockquote>",
                html);
        }
    }
}
