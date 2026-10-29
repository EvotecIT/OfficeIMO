using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_TableStyleFirstRowBecomesStyledHeader() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(3, 2, WordTableStyle.OfficeIMOModern);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Item";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Amount";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Licences";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "12";

            string html = doc.ToHtml();

            Assert.Matches("(?is)<thead>\\s*<tr>\\s*<th[^>]*scope=\"col\"[^>]*>", html);
            Assert.Matches("(?is)<th[^>]*font-weight:bold[^>]*>.*?Item", html);
            Assert.Matches("(?is)<th[^>]*background-color:#EAF0FC[^>]*>", html);
            Assert.Matches("(?is)<th[^>]*border-bottom:1pt solid #3F64D7[^>]*>", html);
            Assert.Matches("(?is)<tbody>.*Licences.*</tbody>", html);
        }

        [Fact]
        public void Test_WordToHtml_GridSpanCellKeepsColumnSpan() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(2, 3);
            for (int c = 0; c < 3; c++) table.Rows[0].Cells[c].Paragraphs[0].Text = "H" + c;
            TableRow lastRow = table.Rows[1]._tableRow;
            TableCell[] cells = lastRow.Elements<TableCell>().ToArray();
            cells[1].Remove();
            TableCellProperties properties = cells[0].TableCellProperties ?? cells[0].PrependChild(new TableCellProperties());
            properties.GridSpan = new GridSpan { Val = 2 };
            cells[0].Elements<Paragraph>().First().AppendChild(new Run(new Text("Total")));
            cells[2].Elements<Paragraph>().First().AppendChild(new Run(new Text("46.75")));

            string html = doc.ToHtml();

            Assert.Matches("(?is)<td[^>]*colspan=\"2\"[^>]*>.*?Total", html);
        }

        [Fact]
        public void Test_WordToHtml_FixedWidthTableUsesShareOfTextColumn() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 2);
            table.WidthType = WordTableWidthUnit.Dxa;
            table.Width = 4680;

            string html = doc.ToHtml();

            Assert.Contains("<table style=\"width:50%", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WordToHtml_DefaultsToTrackedChangeMarkup() {
            Assert.Equal(WordTrackedChangeExportPolicy.Markup, new WordHtmlExportOptions().TrackedChangePolicy);
            Assert.Equal(WordTrackedChangeExportPolicy.Markup, WordHtmlExportOptions.CreateSemanticDocumentProfile().TrackedChangePolicy);
            Assert.Equal(WordTrackedChangeExportPolicy.Markup, WordHtmlExportOptions.CreatePrintReviewProfile().TrackedChangePolicy);
        }

        [Fact]
        public void Test_WordToHtml_UsesDocumentFontsBeforeThemeStack() {
            using var doc = WordDocument.Create();
            doc.Settings.FontFamily = "Calibri";
            doc.AddParagraph("Body");

            string html = doc.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = true });

            Assert.Contains("body{font-family:\"Calibri\", sans-serif}", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WordToHtml_TitleFallsBackToFirstHeading() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Quarterly report").Style = WordParagraphStyles.Heading1;

            string html = doc.ToHtml();

            Assert.Contains("<title>Quarterly report</title>", html, StringComparison.Ordinal);
        }
    }
}

namespace OfficeIMO.Tests {
    public class HtmlToWordCssPixelSizes {
        [Fact]
        public void Test_HtmlToWord_PixelFontSizesConvertToPoints() {
            string html = "<h1 style=\"font-size:32px\">Title</h1><p style=\"font-size:14px\">Body text</p><p style=\"font-size:12pt\">Points</p>";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new WordHtmlImportOptions());

            var title = doc.Paragraphs.First(p => p.Text == "Title");
            var body = doc.Paragraphs.First(p => p.Text == "Body text");
            var points = doc.Paragraphs.First(p => p.Text == "Points");
            Assert.Equal(24D, title.GetRuns().First(r => r.Text.Length > 0).FontSizePoints);
            Assert.Equal(10.5D, body.GetRuns().First(r => r.Text.Length > 0).FontSizePoints);
            Assert.Equal(12D, points.GetRuns().First(r => r.Text.Length > 0).FontSizePoints);
        }
    }
}
