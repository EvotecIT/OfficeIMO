using System;
using System.Collections.Generic;
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
        public void TableBlock_RenderMarkdown_EscapesLiteralBackslashesBeforeMarkdownPunctuation() {
            var table = new TableBlock();
            table.Headers.Add("Header");

            table.Rows.Add(new[] { @"\*\*not bold\*\* and \| literal pipe" });

            var markdown = ((IMarkdownBlock)table).RenderMarkdown();
            MarkdownDoc parsed = MarkdownReader.Parse(markdown);
            TableBlock parsedTable = Assert.IsType<TableBlock>(Assert.Single(parsed.Blocks));

            const string expected = "| Header |\n" +
                                    "| --- |\n" +
                                    @"| \\\*\\\*not bold\\\*\\\* and \\\| literal pipe |";

            Assert.Equal(expected, markdown);
            Assert.Equal(@"\*\*not bold\*\* and \| literal pipe", ExtractPlainText(parsedTable.RowInlines[0][0]));
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

        [Fact]
        public void TableBlock_RenderHtml_PreservesPhysicalCellsBeyondExpandedSpanLimit() {
            var table = new TableBlock();
            var row = new string[TableBlock.MaxEffectiveColumnCount + 1];
            row[TableBlock.MaxEffectiveColumnCount] = "LastPhysicalCell";
            table.Rows.Add(row);

            var html = ((IMarkdownBlock)table).RenderHtml();

            Assert.Contains("LastPhysicalCell", html);
            Assert.Equal(TableBlock.MaxEffectiveColumnCount + 1, table.RowCells[0].Count);
            Assert.Equal(0, table.SkippedColumnCount);
        }

        [Fact]
        public void TableBlock_RenderHtml_ClampsStructuredColSpansToExpandedLimit() {
            var table = new TableBlock();
            var headerCells = new List<TableCell>();
            for (int i = 0; i < 1000; i++) {
                string text = i == 8 ? "Blocked" : $"Visible{i}";
                table.Headers.Add(text);
                headerCells.Add(CreateHeaderCell(text, columnSpan: 512));
            }

            table.SetStructuredCells(headerCells, rows: null, table.ComputeContentSignature());

            var html = ((IMarkdownBlock)table).RenderHtml();

            Assert.Equal(8, CountOccurrences(html, "colspan=\"512\""));
            Assert.Contains("Visible7", html);
            Assert.DoesNotContain("Blocked", html);
        }

        private static TableCell CreateHeaderCell(string text, int columnSpan) {
            return new TableCell(new[] {
                new ParagraphBlock(new InlineSequence().Text(text))
            }) {
                ColumnSpan = columnSpan
            };
        }

        private static int CountOccurrences(string value, string search) {
            int count = 0;
            int index = 0;
            while ((index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0) {
                count++;
                index += search.Length;
            }

            return count;
        }

        private static string ExtractPlainText(IPlainTextMarkdownInline inline) {
            var builder = new System.Text.StringBuilder();
            inline.AppendPlainText(builder);
            return builder.ToString();
        }
    }
}
