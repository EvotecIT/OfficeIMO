using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_TableCellAlignmentAndFormatting() {
            string md = "| Left | Center | Right |\n| :--- | :---: | ---: |\n| **Bold** | *Italic* | Normal |";
            var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions { FontFamily = "Calibri" });
            var table = doc.Tables[0];

            var leftParagraph = table.Rows[1].Cells[0].Paragraphs[0];
            var centerParagraph = table.Rows[1].Cells[1].Paragraphs[0];
            var rightParagraph = table.Rows[1].Cells[2].Paragraphs[0];

            Assert.Equal(JustificationValues.Left, leftParagraph.ParagraphAlignment);
            Assert.Equal(JustificationValues.Center, centerParagraph.ParagraphAlignment);
            Assert.Equal(JustificationValues.Right, rightParagraph.ParagraphAlignment);

            Assert.Contains(leftParagraph.GetRuns(), r => r.Bold);
            Assert.Contains(centerParagraph.GetRuns(), r => r.Italic);
        }

        [Fact]
        public void MarkdownToWord_AppliesSharedVisualThemeToHeadingsAndTables() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithColors(
                    heading: "#064e3b",
                    text: "#102030",
                    border: "#123456",
                    tableHeaderBackground: "#fedcba",
                    tableHeaderText: "#010203",
                    tableStripeBackground: "#f0f9ff")
                .WithTable(table => {
                    table.BorderWidth = 1.25;
                    table.CellPaddingX = 9;
                    table.CellPaddingY = 4;
                    table.UseRowStripes = true;
                });
            string md = """
# Report Theme

Narrative body text.

| Name | Value |
| --- | --- |
| First | 1 |
| Second | 2 |
""";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions { Theme = theme });

            var heading = doc.Paragraphs.First(p => p.Text == "Report Theme");
            Assert.Contains(heading.GetRuns(), run => string.Equals(run.ColorHex, "064e3b", System.StringComparison.OrdinalIgnoreCase));
            var body = doc.Paragraphs.First(p => p.Text == "Narrative body text.");
            Assert.Contains(body.GetRuns(), run => string.Equals(run.ColorHex, "102030", System.StringComparison.OrdinalIgnoreCase));

            var table = doc.Tables[0];
            Assert.Equal("FEDCBA", table.Rows[0].Cells[0].ShadingFillColorHex);
            Assert.Equal("123456", table.Rows[0].Cells[0].Borders.TopColorHex);
            Assert.Equal(10U, table.Rows[0].Cells[0].Borders.TopSize?.Value);
            Assert.Equal((short?)180, table.Rows[0].Cells[0].MarginLeftWidth);
            Assert.Equal((short?)180, table.Rows[0].Cells[0].MarginRightWidth);
            Assert.Equal((short?)80, table.Rows[0].Cells[0].MarginTopWidth);
            Assert.Equal((short?)80, table.Rows[0].Cells[0].MarginBottomWidth);
            Assert.Contains(table.Rows[0].Cells[0].Paragraphs.SelectMany(p => p.GetRuns()), run => string.Equals(run.ColorHex, "010203", System.StringComparison.OrdinalIgnoreCase));
            Assert.Equal("F0F9FF", table.Rows[2].Cells[0].ShadingFillColorHex);
        }

        [Fact]
        public void MarkdownToWord_SharedVisualTheme_PreservesLinkAccentInsideHeadings() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithColors(heading: "#064e3b", accent: "#dc2626");
            string md = "# [Documentation](https://example.com/docs)";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions { Theme = theme });

            var linkRuns = doc.Paragraphs.SelectMany(p => p.GetRuns())
                .Where(run => run.IsHyperLink && run.Text == "Documentation")
                .ToArray();
            Assert.NotEmpty(linkRuns);
            Assert.All(linkRuns, run => Assert.Equal("DC2626", run.ColorHex));
        }

        [Fact]
        public void MarkdownToWord_AppliesDefaultSharedVisualThemeWhenThemeIsOmitted() {
            string md = """
# Heading

Narrative body text.
""";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions());

            var heading = doc.Paragraphs.First(p => p.Text == "Heading");
            Assert.Contains(heading.GetRuns(), run => string.Equals(run.ColorHex, "111827", System.StringComparison.OrdinalIgnoreCase));
            var body = doc.Paragraphs.First(p => p.Text == "Narrative body text.");
            Assert.Contains(body.GetRuns(), run => string.Equals(run.ColorHex, "1f2937", System.StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownToWord_CanDisableDefaultSharedVisualTheme() {
            string md = """
# Heading

Narrative body text.
""";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions { ApplyDefaultTheme = false });

            var heading = doc.Paragraphs.First(p => p.Text == "Heading");
            Assert.DoesNotContain(heading.GetRuns(), run => !string.IsNullOrWhiteSpace(run.ColorHex));
            var body = doc.Paragraphs.First(p => p.Text == "Narrative body text.");
            Assert.DoesNotContain(body.GetRuns(), run => !string.IsNullOrWhiteSpace(run.ColorHex));
        }

        [Fact]
        public void MarkdownToWord_SharedVisualTheme_AllowsBorderlessTables() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithTable(table => table.BorderWidth = 0);
            string md = """
| Name | Value |
| --- | --- |
| First | 1 |
""";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions { Theme = theme });

            var cell = doc.Tables[0].Rows[0].Cells[0];
            Assert.Equal(BorderValues.None, cell.Borders.TopStyle);
            Assert.Equal(0U, cell.Borders.TopSize?.Value);
            Assert.Equal(BorderValues.None, cell.Borders.BottomStyle);
            Assert.Equal(0U, cell.Borders.BottomSize?.Value);
        }

        [Fact]
        public void MarkdownToWord_SharedVisualTheme_TreatsTransparentTableBordersAsNoBorder() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithColors(border: "Transparent")
                .WithTable(table => table.BorderWidth = 1.25);
            string md = """
| Name | Value |
| --- | --- |
| First | 1 |
""";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions { Theme = theme });

            var cell = doc.Tables[0].Rows[0].Cells[0];
            Assert.Equal(BorderValues.None, cell.Borders.TopStyle);
            Assert.Equal(0U, cell.Borders.TopSize?.Value);
            Assert.True(string.IsNullOrEmpty(cell.Borders.TopColorHex));
            Assert.Equal(BorderValues.None, cell.Borders.BottomStyle);
            Assert.Equal(0U, cell.Borders.BottomSize?.Value);
            Assert.True(string.IsNullOrEmpty(cell.Borders.BottomColorHex));
        }

        [Fact]
        public void MarkdownToWord_SharedVisualTheme_TreatsTransparentCodeBackgroundAsNoFill() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithColors(codeBackground: "Transparent");
            string md = """
```csharp
Console.WriteLine("ready");
```
""";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions { Theme = theme });

            var codeParagraph = doc.Paragraphs.First(p => p.Text.Contains("Console.WriteLine", System.StringComparison.Ordinal));
            Assert.True(string.IsNullOrEmpty(codeParagraph.ShadingFillColorHex));
        }

        [Fact]
        public void MarkdownToWord_SharedVisualTheme_TreatsTransparentTableFillsAsNoFill() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithColors(
                    tableHeaderBackground: "Transparent",
                    tableHeaderText: "Transparent",
                    tableStripeBackground: "#00000000")
                .WithTable(table => table.UseRowStripes = true);
            string md = """
| Name | Value |
| --- | --- |
| First | 1 |
| Second | 2 |
""";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions { Theme = theme });

            var table = doc.Tables[0];
            Assert.True(string.IsNullOrEmpty(table.Rows[0].Cells[0].ShadingFillColorHex));
            Assert.DoesNotContain(table.Rows[0].Cells[0].Paragraphs.SelectMany(p => p.GetRuns()), run => string.Equals(run.ColorHex, "000000", System.StringComparison.OrdinalIgnoreCase));
            Assert.True(string.IsNullOrEmpty(table.Rows[2].Cells[0].ShadingFillColorHex));
        }

        [Fact]
        public void MarkdownToWord_SharedVisualTheme_TreatsTransparentCalloutTitleSurfaceAsNoFill() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithColors(surface: "Transparent");
            string md = """
> [!NOTE] Heads up
> Body text.
""";

            using var doc = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToWordDocument(new MarkdownToWordOptions { Theme = theme });

            var title = doc.Paragraphs.First(p => p.Text == "Heads up");
            Assert.True(string.IsNullOrEmpty(title.ShadingFillColorHex));
        }

        [Fact]
        public void WordToMarkdown_TableAlignmentMarkers() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(2, 3);

            var left = table.Rows[0].Cells[0].Paragraphs[0];
            left.Text = "Left";
            left.ParagraphAlignment = JustificationValues.Left;

            var center = table.Rows[0].Cells[1].Paragraphs[0];
            center.Text = "Center";
            center.ParagraphAlignment = JustificationValues.Center;

            var right = table.Rows[0].Cells[2].Paragraphs[0];
            right.Text = "Right";
            right.ParagraphAlignment = JustificationValues.Right;

            table.Rows[1].Cells[0].Paragraphs[0].Text = "A";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "B";
            table.Rows[1].Cells[2].Paragraphs[0].Text = "C";

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());

            var lines = markdown.Split('\n');
            Assert.Equal("| Left | Center | Right |", lines[0].TrimEnd('\r'));
            Assert.Equal("| :--- | :---: | ---: |", lines[1].TrimEnd('\r'));
        }

        [Fact]
        public void MarkdownDoc_ToWordDocument_Preserves_MultiParagraph_TableCell_From_HtmlAst() {
            const string html = """
                <table>
                  <tr>
                    <th>Status</th>
                  </tr>
                  <tr>
                    <td>
                      <p>Healthy</p>
                      <p>Observed from AST path</p>
                    </td>
                  </tr>
                </table>
                """;

            var markdown = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToMarkdownDocument();

            using var doc = markdown.ToWordDocument(new MarkdownToWordOptions { FontFamily = "Calibri" });
            var cellParagraphs = doc.Tables[0].Rows[1].Cells[0].Paragraphs
                .Select(p => p.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            Assert.Equal(2, cellParagraphs.Count);
            Assert.Equal("Healthy", cellParagraphs[0]);
            Assert.Equal("Observed from AST path", cellParagraphs[1]);
        }

        [Fact]
        public void Html_ToWordDocumentViaMarkdown_Preserves_MultiParagraph_TableCell_From_AstBridge() {
            const string html = """
                <table>
                  <tr>
                    <th>Status</th>
                  </tr>
                  <tr>
                    <td>
                      <p>Healthy</p>
                      <p>Observed from bridge</p>
                    </td>
                  </tr>
                </table>
                """;

            MarkdownDoc markdown = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToMarkdownDocument();
            using var doc = markdown.ToWordDocument(new MarkdownToWordOptions { FontFamily = "Calibri" });
            var cellParagraphs = doc.Tables[0].Rows[1].Cells[0].Paragraphs
                .Select(p => p.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            Assert.Equal(2, cellParagraphs.Count);
            Assert.Equal("Healthy", cellParagraphs[0]);
            Assert.Equal("Observed from bridge", cellParagraphs[1]);
        }

        [Fact]
        public void MarkdownDoc_ToWordDocument_Bridges_RawHtml_TableCell_Content_Through_Ast() {
            const string html = """
                <table>
                  <tr>
                    <th>Notes</th>
                  </tr>
                  <tr>
                    <td>
                      <custom-card>
                        <p>Alpha</p>
                        <p><strong>Beta</strong></p>
                      </custom-card>
                    </td>
                  </tr>
                </table>
                """;

            var markdown = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToMarkdownDocument();
            var rawCellBlock = Assert.IsType<OfficeIMO.Markdown.HtmlRawBlock>(Assert.Single(Assert.Single(markdown.Blocks.OfType<OfficeIMO.Markdown.TableBlock>()).RowCells[0][0].ChildBlocks));

            Assert.Contains("<custom-card>", rawCellBlock.Html);

            using var doc = markdown.ToWordDocument(new MarkdownToWordOptions { FontFamily = "Calibri" });
            var cellParagraphs = doc.Tables[0].Rows[1].Cells[0].Paragraphs
                .Select(p => p.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            Assert.Equal(2, cellParagraphs.Count);
            Assert.Equal("Alpha", cellParagraphs[0]);
            Assert.Equal("Beta", cellParagraphs[1]);
            Assert.Contains(doc.Tables[0].Rows[1].Cells[0].Paragraphs.SelectMany(p => p.GetRuns()), run => run.Bold && string.Equals(run.Text, "Beta", System.StringComparison.Ordinal));
        }
    }
}
