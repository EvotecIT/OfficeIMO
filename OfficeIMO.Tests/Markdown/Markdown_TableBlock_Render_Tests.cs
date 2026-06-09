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
            table.Rows.Add(new[] { "Multi\r\nLine", "Pipe|And\\Back" });

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
            table.Rows.Add(new[] { "Multi\r\nLine", "Pipe|And\\Back" });

            var html = ((IMarkdownBlock)table).RenderHtml();

            const string expected = "<table><thead><tr><th>Name|Title</th><th>Path \\ Server</th></tr></thead><tbody>" +
                                    "<tr><td>Cell | one</td><td>C: \\ Share</td></tr>" +
                                    "<tr><td>Multi<br/>Line</td><td>Pipe|And\\Back</td></tr>" +
                                    "</tbody></table>";

            Assert.Equal(expected, html);
        }

        [Fact]
        public void TableBlock_RenderMarkdown_PreservesExistingBreakTags() {
            var table = new TableBlock();
            table.Headers.Add("Header");

            table.Rows.Add(new[] { "Line1<br/>Line2" });

            var markdown = ((IMarkdownBlock)table).RenderMarkdown();

            const string expected = "| Header |\n" +
                                    "| --- |\n" +
                                    "| Line1<br/>Line2 |";

            Assert.Equal(expected, markdown);
        }

        [Fact]
        public void TableBlock_RenderHtml_PreservesExistingBreakTags() {
            var table = new TableBlock();
            table.Headers.Add("Header");

            table.Rows.Add(new[] { "Line1<br/>Line2" });

            var html = ((IMarkdownBlock)table).RenderHtml();

            const string expected = "<table><thead><tr><th>Header</th></tr></thead><tbody>" +
                                    "<tr><td>Line1<br/>Line2</td></tr>" +
                                    "</tbody></table>";

            Assert.Equal(expected, html);
        }

        [Fact]
        public void TableBlock_RenderMarkdown_PadsRowsToHeaderCount() {
            var table = new TableBlock();
            table.Headers.Add("Col1");
            table.Headers.Add("Col2");

            table.Rows.Add(new[] { "Value" });

            var markdown = ((IMarkdownBlock)table).RenderMarkdown();

            const string expected = "| Col1 | Col2 |\n" +
                                    "| --- | --- |\n" +
                                    "| Value |  |";

            Assert.Equal(expected, markdown);
        }

        [Fact]
        public void TableBlock_RenderHtml_PadsRowsToHeaderCount() {
            var table = new TableBlock();
            table.Headers.Add("Col1");
            table.Headers.Add("Col2");

            table.Rows.Add(new[] { "Value" });

            var html = ((IMarkdownBlock)table).RenderHtml();

            const string expected = "<table><thead><tr><th>Col1</th><th>Col2</th></tr></thead><tbody>" +
                                    "<tr><td>Value</td><td></td></tr>" +
                                    "</tbody></table>";

            Assert.Equal(expected, html);
        }

        [Fact]
        public void TableBlock_RenderHtml_EmitsColumnWidthHints() {
            var table = new TableBlock();
            table.Headers.Add("Code");
            table.Headers.Add("Description");
            table.ColumnWidthPoints.Add(54D);
            table.ColumnWidthPoints.Add(null);
            table.ColumnWidthWeights.Add(1D);
            table.ColumnWidthWeights.Add(3D);

            table.Rows.Add(new[] { "A-100", "Consulting" });

            var html = ((IMarkdownBlock)table).RenderHtml();

            const string expected = "<table><colgroup><col style=\"width:54pt\"><col style=\"width:75%\"></colgroup><thead><tr><th>Code</th><th>Description</th></tr></thead><tbody>" +
                                    "<tr><td>A-100</td><td>Consulting</td></tr>" +
                                    "</tbody></table>";

            Assert.Equal(expected, html);
        }

        [Fact]
        public void TableBlock_RenderHtml_SanitizesDisallowedTags() {
            var table = new TableBlock();
            table.Headers.Add("Header");

            table.Rows.Add(new[] { "<script>alert(1)</script>" });

            var html = ((IMarkdownBlock)table).RenderHtml();

            const string expected = "<table><thead><tr><th>Header</th></tr></thead><tbody>" +
                                    "<tr><td>&lt;script&gt;alert(1)&lt;/script&gt;</td></tr>" +
                                    "</tbody></table>";

            Assert.Equal(expected, html);
        }
    }
}
