using AngleSharp.Html.Parser;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_TableOfContentsIsRegeneratedFromHeadings() {
            using WordDocument document = WordDocument.Create();
            document.AddTableOfContent();
            document.AddParagraph("Installation").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Body text.");
            document.AddParagraph("Configuration").Style = WordParagraphStyles.Heading1;
            document.AddParagraph("Environment variables").Style = WordParagraphStyles.Heading2;

            var result = document.ConvertToHtmlWithEvidence(new WordHtmlExportOptions());
            var html = new HtmlParser().ParseDocument(result.RequireValue());

            var nav = Assert.Single(html.QuerySelectorAll("nav.toc"));
            var links = nav.QuerySelectorAll("a").ToArray();
            Assert.Equal(new[] { "Installation", "Configuration", "Environment variables" }, links.Select(link => link.TextContent).ToArray());
            foreach (var link in links) {
                string href = link.GetAttribute("href")!;
                Assert.StartsWith("#", href);
                var target = html.GetElementById(href.Substring(1));
                Assert.NotNull(target);
                Assert.Equal(link.TextContent, target!.TextContent.Trim());
            }
            // The level-2 entry nests under its level-1 parent.
            Assert.NotNull(nav.QuerySelector("li li a"));
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "toc-regenerated-from-headings");
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_WordToHtml_DefaultCssPreservesSourceTableFormatting(bool includeCss) {
            using WordDocument document = WordDocument.Create();
            var table = document.AddTable(3, 2, WordTableStyle.TableNormal);
            table.RepeatHeaderRowAtTheTopOfEachPage = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Revenue";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "EMEA";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "1,250";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "APAC";
            table.Rows[2].Cells[1].Paragraphs[0].Text = "-12.5%";
            table.Rows[2].Cells[1].Paragraphs[0].ParagraphAlignment = WordParagraphAlignment.Right;

            string output = document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = includeCss });
            var html = new HtmlParser().ParseDocument(output);

            if (includeCss) Assert.Contains("--office-accent-1: #" + OfficeTheme.Default.Palette.Accent1.ToRgbHex(), output);
            var rows = html.QuerySelectorAll("tr").ToArray();
            Assert.All(rows[0].Children, cell => {
                if (!includeCss) Assert.Contains("font-weight:normal", cell.GetAttribute("style"));
                Assert.Contains("text-align:left", cell.GetAttribute("style"));
            });
            Assert.DoesNotContain("text-align:right", rows[1].Children[0].GetAttribute("style") ?? string.Empty);
            Assert.DoesNotContain("text-align:right", rows[1].Children[1].GetAttribute("style") ?? string.Empty);
            Assert.Contains("text-align:right", rows[2].Children[1].GetAttribute("style") ?? string.Empty);
            var cssDocument = OfficeIMO.Html.HtmlConversionDocument.Parse(output);
            var computed = OfficeIMO.Html.HtmlComputedStyleEngine.Compute(cssDocument);
            Assert.All(cssDocument.Document.QuerySelectorAll("td"), cell => {
                Assert.Empty(computed[cell].GetValue("background-color"));
                Assert.Empty(computed[cell].GetValue("border-bottom-style"));
                Assert.Empty(computed[cell].GetValue("padding-top"));
            });
        }

        [Fact]
        public void Test_WordToHtml_MonospaceRunsBecomeInlineCodeAndCellsStayInline() {
            using WordDocument document = WordDocument.Create();
            var paragraph = document.AddParagraph("Settings live in ");
            paragraph.AddText("lumen.json").SetFontFamily("Consolas");
            paragraph.AddText(" at the root.");
            var table = document.AddTable(1, 2);
            table.RepeatHeaderRowAtTheTopOfEachPage = false;
            table.Rows[0].Cells[0].Paragraphs[0].AddText("title").SetFontFamily("Consolas");
            table.Rows[0].Cells[1].Paragraphs[0].AddText("string");
            var block = document.AddParagraph(string.Empty);
            block.AddText(string.Empty).Text = "line one\n    line two";
            foreach (var run in block.GetRuns()) run.SetFontFamily("Courier New");

            var html = new HtmlParser().ParseDocument(document.ToHtml());

            var body = html.Body!;
            Assert.Contains(body.QuerySelectorAll("p code"), code => code.TextContent == "lumen.json");
            Assert.Contains(body.QuerySelectorAll("td code"), code => code.TextContent == "title");
            Assert.Empty(body.QuerySelectorAll("td pre"));
            var pre = Assert.Single(body.QuerySelectorAll("pre code"));
            Assert.Equal("line one\n    line two", pre.TextContent);
        }
    }
}
