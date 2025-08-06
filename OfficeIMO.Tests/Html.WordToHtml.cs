using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_HeadingsAndFormatting() {
            using var doc = WordDocument.Create();
            doc.BuiltinDocumentProperties.Title = "Test Document";

            var h1 = doc.AddParagraph("Heading 1");
            h1.Style = WordParagraphStyles.Heading1;

            var p = doc.AddParagraph();
            p.AddText("bold").Bold = true;
            p.AddText(" and ");
            p.AddText("italic").Italic = true;
            p.AddText(" and ");
            p.AddText("underline").Underline = UnderlineValues.Single;

            var link = doc.AddParagraph();
            link.AddHyperLink("GitHub", new Uri("https://github.com"));

            string html = doc.ToHtml();

            Assert.Contains("<h1>Heading 1</h1>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<strong>bold</strong>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<em>italic</em>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<u>underline</u>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("https://github.com", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<title>Test Document</title>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_ListsAndTable() {
            using var doc = WordDocument.Create();

            var list = doc.AddList(WordListStyle.Bulleted);
            list.AddItem("Item 1");
            list.AddItem("Sub 1", 1);
            list.AddItem("Item 2");

            var ordered = doc.AddList(WordListStyle.ArticleSections);
            ordered.AddItem("First");
            ordered.AddItem("Second");

            var table = doc.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "A";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "B";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "C";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "D";

            string html = doc.ToHtml();

            Assert.Contains("<ul", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<ol", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Sub 1", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<table>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">A<", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">D<", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_ImageAndMetadata() {
            using var doc = WordDocument.Create();
            doc.BuiltinDocumentProperties.Title = "With Image";
            doc.BuiltinDocumentProperties.Creator = "Tester";

            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            doc.AddParagraph().AddImage(assetPath);

            string html = doc.ToHtml();

            Assert.Contains("data:image/png", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("content=\"Tester\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_ListStartAttribute() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.Headings111);
            list.Numbering.Levels[0].SetStartNumberingValue(4);
            list.AddItem("Four");
            list.AddItem("Five");

            string html = doc.ToHtml();

            Assert.Contains("<ol start=\"4\" type=\"1\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_RomanNumerals() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.HeadingIA1);
            list.AddItem("Intro");
            list.AddItem("Body");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });

            Assert.Contains("<ol type=\"I\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list-style-type:upper-roman", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_BulletType() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.Bulleted);
            list.Numbering.Levels[0]._level.LevelText.Val = "o";
            list.AddItem("One");
            list.AddItem("Two");

            string html = doc.ToHtml();

            Assert.Contains("<ul type=\"circle\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_TableCellStyles() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 1);
            table.WidthType = TableWidthUnitValues.Pct;
            table.Width = 5000;

            var cell = table.Rows[0].Cells[0];
            cell.WidthType = TableWidthUnitValues.Pct;
            cell.Width = 2500;
            cell.Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
            cell.Borders.LeftStyle = BorderValues.Single;
            cell.Borders.RightStyle = BorderValues.Single;
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.BottomStyle = BorderValues.Single;

            string html = doc.ToHtml();

            Assert.Contains("<table style=\"width:100%;border:1px solid black;border-collapse:collapse\">", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<td style=\"width:50%;text-align:center;border:1px solid black\">", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_Blockquote() {
            using var doc = WordDocument.Create();
            var p = doc.AddParagraph("Quoted text");
            p.IndentationBefore = 720;

            string html = doc.ToHtml();

            Assert.Contains("<blockquote>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Quoted text", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_HorizontalRule() {
            using var doc = WordDocument.Create();
            doc.AddHorizontalLine();

            string html = doc.ToHtml();

            Assert.Contains("<hr", html, StringComparison.OrdinalIgnoreCase);
        }
    }
}

