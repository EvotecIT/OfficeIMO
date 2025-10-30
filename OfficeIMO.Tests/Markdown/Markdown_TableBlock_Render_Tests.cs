using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_TableBlock_Render_Tests {
        [Fact]
        public void TableBlock_RenderMarkdown_EscapesSpecialCharacters() {
            var table = new TableBlock();
            table.Headers.Add("Name|Title");
            table.Headers.Add("Path \\ Server");

            table.Rows.Add(new[] { "Cell | one", "C: \\ Share" });
            table.Rows.Add(new[] { "Multi\nLine", "Pipe|And\\Back" });

            var markdown = ((IMarkdownBlock)table).RenderMarkdown();

            const string expected = "| Name\\|Title | Path \\\\ Server |\n" +
                                    "| --- | --- |\n" +
                                    "| Cell \\| one | C: \\\\ Share |\n" +
                                    "| Multi<br>Line | Pipe\\|And\\\\Back |";

            Assert.Equal(expected, markdown);
        }

        [Fact]
        public void TableBlock_RenderHtml_PreservesPipesAndBackslashes() {
            var table = new TableBlock();
            table.Headers.Add("Name|Title");
            table.Headers.Add("Path \\ Server");

            table.Rows.Add(new[] { "Cell | one", "C: \\ Share" });
            table.Rows.Add(new[] { "Multi\nLine", "Pipe|And\\Back" });

            var html = ((IMarkdownBlock)table).RenderHtml();

            const string expected = "<table><thead><tr><th>Name|Title</th><th>Path \\ Server</th></tr></thead><tbody>" +
                                    "<tr><td>Cell | one</td><td>C: \\ Share</td></tr>" +
                                    "<tr><td>Multi<br/>Line</td><td>Pipe|And\\Back</td></tr>" +
                                    "</tbody></table>";

            Assert.Equal(expected, html);
        }
    }
}
