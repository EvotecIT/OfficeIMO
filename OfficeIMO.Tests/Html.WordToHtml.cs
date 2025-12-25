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
        public void Test_WordToHtml_AnchorLinks() {
            using var doc = WordDocument.Create();
            var target = doc.AddParagraph("Target");
            WordBookmark.AddBookmark(target, "anchor");
            var link = doc.AddParagraph();
            link.AddHyperLink("Go", "anchor");

            string html = doc.ToHtml();

            Assert.Contains("href=\"#anchor\"", html, StringComparison.OrdinalIgnoreCase);
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
            var paragraph = doc.AddParagraph();
            paragraph.AddImage(assetPath, 40, 40, description: "Company logo");

            Assert.NotNull(paragraph.Image);
            Assert.Equal("Company logo", paragraph.Image!.Description);

            string html = doc.ToHtml();

            Assert.Contains("data:image/png", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("content=\"Tester\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("alt=\"Company logo\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("width=\"40\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("height=\"40\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_ImageFilePathOption() {
            using var doc = WordDocument.Create();

            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            var paragraph = doc.AddParagraph();
            paragraph.AddImage(assetPath, 20, 20);

            string html = doc.ToHtml(new WordToHtmlOptions { EmbedImagesAsBase64 = false });

            Assert.DoesNotContain("data:image", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(Path.GetFileName(assetPath), html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_ListStartAttribute() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.Numbered);
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

            Assert.Contains("<ol start=\"1\" type=\"I\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list-style-type:upper-roman", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_BulletType() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.Bulleted);
            list.Numbering.Levels[0]._level.LevelText!.Val = "o";
            list.AddItem("One");
            list.AddItem("Two");

            string html = doc.ToHtml();

            Assert.Contains("<ul type=\"circle\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_LowerLetter() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.ArticleSections);
            list.Numbering.Levels[0]._level.NumberingFormat!.Val = NumberFormatValues.LowerLetter;
            list.AddItem("alpha");
            list.AddItem("beta");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });

            Assert.Contains("<ol start=\"1\" type=\"a\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list-style-type:lower-alpha", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_DecimalLeadingZero() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.ArticleSections);
            list.Numbering.Levels[0]._level.NumberingFormat!.Val = NumberFormatValues.DecimalZero;
            list.Numbering.Levels[0].SetStartNumberingValue(3);
            list.AddItem("three");
            list.AddItem("four");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });

            Assert.Contains("<ol", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("start=\"3\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list-style-type:decimal-leading-zero", html, StringComparison.OrdinalIgnoreCase);
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
        public void Test_WordToHtml_TableCellCss() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 1);
            var cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
            cell.ShadingFillColorHex = "ff0000";
            cell.Borders.LeftStyle = BorderValues.Single;
            cell.Borders.RightStyle = BorderValues.Single;
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.BottomStyle = BorderValues.Single;
            cell.Borders.LeftColorHex = "00ff00";
            cell.Borders.RightColorHex = "00ff00";
            cell.Borders.TopColorHex = "00ff00";
            cell.Borders.BottomColorHex = "00ff00";
            cell.Borders.LeftSize = 8;
            cell.Borders.RightSize = 8;
            cell.Borders.TopSize = 8;
            cell.Borders.BottomSize = 8;

            string html = doc.ToHtml();

            Assert.Contains("background-color:#ff0000", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-align:right", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("border:1px solid #00ff00", html, StringComparison.OrdinalIgnoreCase);
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

        [Fact]
        public void Test_WordToHtml_AdditionalHeadElements() {
            using var doc = WordDocument.Create();
            doc.BuiltinDocumentProperties.Creator = "Tester";
            doc.AddParagraph("Content");

            var options = new WordToHtmlOptions();
            options.AdditionalMetaTags.Add(("viewport", "width=device-width"));
            options.AdditionalLinkTags.Add(("stylesheet", "styles.css"));

            string html = doc.ToHtml(options);

            Assert.Contains("<meta name=\"viewport\" content=\"width=device-width\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<link rel=\"stylesheet\" href=\"styles.css\"", html, StringComparison.OrdinalIgnoreCase);
            int authorIndex = html.IndexOf("name=\"author\"", StringComparison.OrdinalIgnoreCase);
            int viewportIndex = html.IndexOf("name=\"viewport\"", StringComparison.OrdinalIgnoreCase);
            Assert.True(viewportIndex > authorIndex);
        }

        [Fact]
        public void Test_WordToHtml_StyleClasses() {
            using var doc = WordDocument.Create();
            var p = doc.AddParagraph("Heading with style");
            p.Style = WordParagraphStyles.Heading1;
            p.AddText(" run").CharacterStyleId = "Heading1Char";

            var options = new WordToHtmlOptions { IncludeParagraphClasses = true, IncludeRunClasses = true };
            string html = doc.ToHtml(options);

            Assert.Contains("<h1 class=\"Heading1\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<span class=\"Heading1Char\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(".Heading1 {", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-size:16pt", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(".Heading1Char {", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("color:#2f5496", html, StringComparison.OrdinalIgnoreCase);

            string htmlNoClasses = doc.ToHtml(new WordToHtmlOptions());
            Assert.DoesNotContain("class=\"Heading1\"", htmlNoClasses, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Heading1Char", htmlNoClasses, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_RunColorAndHighlightStyles() {
            using var doc = WordDocument.Create();
            var p = doc.AddParagraph();
            var run = p.AddText("Colored");
            run.ColorHex = "ff0000";
            run.Highlight = HighlightColorValues.Cyan;

            string html = doc.ToHtml(new WordToHtmlOptions {
                IncludeRunColorStyles = true,
                IncludeRunHighlightStyles = true
            });

            Assert.Contains("color:#ff0000", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("background-color:#00ffff", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_ParagraphSpacingAndIndentationStyles() {
            using var doc = WordDocument.Create();
            var p = doc.AddParagraph("Spacing");
            p.LineSpacingBefore = 240;
            p.LineSpacingAfter = 480;
            p.LineSpacing = 360;
            p.LineSpacingRule = LineSpacingRuleValues.Auto;
            p.IndentationAfter = 720;
            p.IndentationFirstLine = 360;

            string html = doc.ToHtml(new WordToHtmlOptions {
                IncludeParagraphSpacingStyles = true,
                IncludeParagraphIndentationStyles = true
            });

            Assert.Contains("margin-top:12pt", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("margin-bottom:24pt", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("line-height:1.5", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("margin-right:36pt", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-indent:18pt", html, StringComparison.OrdinalIgnoreCase);
        }
    }
}

