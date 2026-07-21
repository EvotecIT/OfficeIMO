using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_ExposesSharedTextResultContract() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Result contract");

            HtmlTextConversionResult result = doc.ToHtmlResult();

            Assert.True(result.Succeeded);
            Assert.Contains("Result contract", result.RequireValue(), StringComparison.Ordinal);
        }

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
            Assert.Contains("<p id=\"anchor\">Target</p>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_RunLanguage_RoundTripsLangAttributes() {
            const string sourceHtml = "<html lang=\"en-US\"><body><p>Hello <span lang=\"pl-PL\">Czesc</span></p></body></html>";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(sourceHtml).ToWordDocument();

            Assert.Equal("en-US", doc.Settings.Language);
            Assert.Contains(doc.Paragraphs[0].GetRuns(), run => string.Equals(run.Text, "Czesc", StringComparison.Ordinal) && string.Equals(run.Language, "pl-PL", StringComparison.OrdinalIgnoreCase));

            string html = doc.ToHtml();

            Assert.Contains("<html lang=\"en-US\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<span lang=\"pl-PL\">Czesc</span>", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("lang=\"en-US\">Hello", html, StringComparison.OrdinalIgnoreCase);

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            Assert.Contains(roundTrip.Paragraphs[0].GetRuns(), run => string.Equals(run.Text, "Czesc", StringComparison.Ordinal) && string.Equals(run.Language, "pl-PL", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_WordToHtml_BookmarkedHeading_ExportsId() {
            using var doc = WordDocument.Create();
            var heading = doc.AddParagraph("Bookmarked heading");
            heading.Style = WordParagraphStyles.Heading2;
            heading.AddBookmark("heading-anchor");

            string html = doc.ToHtml();

            Assert.Contains("<h2 id=\"heading-anchor\">Bookmarked heading</h2>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_BookmarkedListItem_ExportsId() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.Bulleted);
            var item = list.AddItem("Bookmarked item");
            item.AddBookmark("item-anchor");

            string html = doc.ToHtml();

            Assert.Contains("<li id=\"item-anchor\">Bookmarked item</li>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_StructuralBookmark_StillExportsStructuralElement() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Article content");
            paragraph.AddBookmark("article:article-anchor");

            string html = doc.ToHtml();

            Assert.Contains("<article id=\"article-anchor\"><p>Article content</p></article>", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<p id=\"article:article-anchor\"", html, StringComparison.OrdinalIgnoreCase);
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
            Assert.Equal(2, html.Split(new[] { "<ul" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(1, html.Split(new[] { "<ol" }, StringSplitOptions.None).Length - 1);
            Assert.Contains("Item 1<ul", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<table>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">A<", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(">D<", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_ListsInsideTable_PreserveStructureAndNesting() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 2);

            var bulletCell = table.Rows[0].Cells[0];
            bulletCell.Paragraphs[0].Remove();
            var bullets = bulletCell.AddList(WordListStyle.Bulleted);
            bullets.AddItem("Bullet root");
            bullets.AddItem("Nested bullet", 1);
            bullets.AddItem("Bullet next");

            var orderedCell = table.Rows[0].Cells[1];
            orderedCell.Paragraphs[0].Remove();
            var numbers = orderedCell.AddList(WordListStyle.Numbered);
            numbers.SetStartNumberingValue(4);
            numbers.AddItem("Number four");
            numbers.AddItem("Number five");

            string html = doc.ToHtml();
            var parsed = HtmlDocumentParser.ParseDocument(html);
            var cells = parsed.QuerySelectorAll("td");
            Assert.Equal(2, cells.Length);

            IElement bulletList = Assert.IsAssignableFrom<IElement>(cells[0].QuerySelector("ul"));
            var bulletItems = bulletList.Children.Where(element => element.LocalName == "li").ToArray();
            Assert.Equal(2, bulletItems.Length);
            Assert.Equal("Bullet rootNested bullet", bulletItems[0].TextContent);
            Assert.Equal("Nested bullet", bulletItems[0].QuerySelector("ul > li")!.TextContent);
            Assert.Equal("Bullet next", bulletItems[1].TextContent);

            IElement orderedList = Assert.IsAssignableFrom<IElement>(cells[1].QuerySelector("ol"));
            Assert.Equal("4", orderedList.GetAttribute("start"));
            Assert.Equal(
                new[] { "Number four", "Number five" },
                orderedList.Children.Where(element => element.LocalName == "li").Select(element => element.TextContent).ToArray());
        }

        [Fact]
        public void Test_WordToHtml_TableHeaderRows_ExportTheadAndTh() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(2, 2);
            table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Score";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Ada";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "42";

            string html = doc.ToHtml();

            Assert.Contains("<thead><tr><th", html, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(2, html.Split(new[] { "scope=\"col\"" }, StringSplitOptions.None).Length - 1);
            Assert.DoesNotContain("<td scope=\"col\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Name</p>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Score</p>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("</th></tr></thead>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<tbody><tr><td", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Ada</p>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>42</p>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("</td></tr></tbody>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_TableHeaderRows_RoundTripMultipleRepeatedRows() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(3, 1);
            table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
            table.Rows[1].RepeatHeaderRowAtTheTopOfEachPage = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Header 1";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Header 2";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "Body";

            string html = doc.ToHtml();

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            var roundTripTable = roundTrip.Tables[0];
            Assert.True(roundTripTable.Rows[0].RepeatHeaderRowAtTheTopOfEachPage);
            Assert.True(roundTripTable.Rows[1].RepeatHeaderRowAtTheTopOfEachPage);
            Assert.False(roundTripTable.Rows[2].RepeatHeaderRowAtTheTopOfEachPage);
        }

        [Fact]
        public void Test_WordToHtml_TableFooterRow_ExportsTfootAndRoundTrips() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(3, 1);
            table.ConditionalFormattingLastRow = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Header";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Body";
            table.Rows[2].Cells[0].Paragraphs[0].Text = "Total";

            string html = doc.ToHtml();

            Assert.Contains("<tbody><tr><td", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Body</p>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<tfoot><tr><td", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Total</p>", html, StringComparison.OrdinalIgnoreCase);
            Assert.True(html.IndexOf("<tbody>", StringComparison.OrdinalIgnoreCase) < html.IndexOf("<tfoot>", StringComparison.OrdinalIgnoreCase));

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            Assert.True(roundTrip.Tables[0].ConditionalFormattingLastRow);
        }

        [Fact]
        public void Test_WordToHtml_TableCaption_ExportsCaptionElementAndRoundTrips() {
            const string sourceHtml = "<table><caption>Score summary</caption><tr><td>Ada</td><td>42</td></tr></table>";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(sourceHtml).ToWordDocument();

            Assert.Contains(doc.Paragraphs, paragraph =>
                string.Equals(paragraph.StyleId, "Caption", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(paragraph.Text, "Score summary", StringComparison.Ordinal));

            string html = doc.ToHtml();

            Assert.Contains("<caption>Score summary</caption>", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<p>Score summary</p><table", html, StringComparison.OrdinalIgnoreCase);

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            Assert.Single(roundTrip.Tables);
            Assert.Contains(roundTrip.Paragraphs, paragraph =>
                string.Equals(paragraph.StyleId, "Caption", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(paragraph.Text, "Score summary", StringComparison.Ordinal));
        }

        [Fact]
        public void Test_WordToHtml_TableCaptionBelow_ExportsCaptionElementWithoutDuplicateParagraph() {
            const string sourceHtml = "<table><caption>Totals</caption><tr><td>42</td></tr></table>";
            var options = new HtmlToWordOptions {
                TableCaptionPosition = TableCaptionPosition.Below
            };

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(sourceHtml).ToWordDocument(options);

            string html = doc.ToHtml();

            Assert.Contains("<caption>Totals</caption>", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("</table><p>Totals</p>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_TopLevelTable_DoesNotExportPlaceholderParagraph() {
            const string sourceHtml = "<table><tr><td>42</td></tr></table>";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(sourceHtml).ToWordDocument();

            string html = doc.ToHtml();

            Assert.Contains("<body><table>", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<body><p></p><table>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_TableWithoutHeaderRows_KeepsFlatRows() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Cell";

            string html = doc.ToHtml();

            Assert.Contains("<table><tr><td", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Cell</p>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("</td></tr></table>", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<thead>", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<tbody>", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<colgroup>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_TableColumnGroups_ExportWhenEnabled() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 2);
            table.WidthType = TableWidthUnitValues.Pct;
            table.Width = 5000;
            table.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Pct;
            table.Rows[0].Cells[0].Width = 675;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "A";
            table.Rows[0].Cells[1].WidthType = TableWidthUnitValues.Pct;
            table.Rows[0].Cells[1].Width = 4325;
            table.Rows[0].Cells[1].Paragraphs[0].Text = "B";

            string defaultHtml = doc.ToHtml();
            Assert.DoesNotContain("<colgroup>", defaultHtml, StringComparison.OrdinalIgnoreCase);

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeTableColumnGroups = true });

            Assert.Contains("<table style=\"width:100%\"><colgroup>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<col style=\"width:13.5%\">", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<col style=\"width:86.5%\">", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("</colgroup><tr><td style=\"width:13.5%\"><p>A</p></td><td style=\"width:86.5%\"><p>B</p></td></tr></table>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_TableColumnGroups_FallBackToGridWidths() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 2);
            table.GridColumnWidth = new List<int> { 2400, 4800 };
            table.Rows[0].Cells[0].WidthType = null;
            table.Rows[0].Cells[0].Width = null;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "A";
            table.Rows[0].Cells[1].WidthType = null;
            table.Rows[0].Cells[1].Width = null;
            table.Rows[0].Cells[1].Paragraphs[0].Text = "B";

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeTableColumnGroups = true });

            Assert.Contains("<colgroup>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<col style=\"width:160px\">", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<col style=\"width:320px\">", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<td style=\"width:", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_HeadersFooters_AreOptIn() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Body text");
            var section = doc.Sections[0];
            section.GetOrCreateHeader(HeaderFooterValues.Default).AddParagraph("Header text");
            section.GetOrCreateFooter(HeaderFooterValues.Default).AddParagraph("Footer text");

            string defaultHtml = doc.ToHtml();

            Assert.DoesNotContain("word-header", defaultHtml, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("word-footer", defaultHtml, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Header text", defaultHtml, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Footer text", defaultHtml, StringComparison.OrdinalIgnoreCase);

            string html = doc.ToHtml(new WordToHtmlOptions { ExportHeadersAndFooters = true });

            Assert.Contains("<header class=\"word-header word-header-default\" data-section-index=\"0\" data-type=\"default\"><p>Header text</p></header>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Body text</p>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<footer class=\"word-footer word-footer-default\" data-section-index=\"0\" data-type=\"default\"><p>Footer text</p></footer>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_HeadersFooters_RoundTripToNativeSections() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Body text");
            var section = doc.Sections[0];
            section.GetOrCreateHeader(HeaderFooterValues.Default).AddParagraph("Header text");
            section.GetOrCreateFooter(HeaderFooterValues.Default).AddParagraph("Footer text");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportHeadersAndFooters = true });

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            Assert.Contains(roundTrip.Paragraphs, paragraph => string.Equals(paragraph.Text, "Body text", StringComparison.Ordinal));
            Assert.DoesNotContain(roundTrip.Paragraphs, paragraph => string.Equals(paragraph.Text, "Header text", StringComparison.Ordinal));
            Assert.DoesNotContain(roundTrip.Paragraphs, paragraph => string.Equals(paragraph.Text, "Footer text", StringComparison.Ordinal));

            var roundTripSection = roundTrip.Sections[0];
            Assert.NotNull(roundTripSection.Header.Default);
            Assert.NotNull(roundTripSection.Footer.Default);
            var header = roundTripSection.Header.Default!;
            var footer = roundTripSection.Footer.Default!;
            Assert.Contains(header.Paragraphs, paragraph => string.Equals(paragraph.Text, "Header text", StringComparison.Ordinal));
            Assert.Contains(footer.Paragraphs, paragraph => string.Equals(paragraph.Text, "Footer text", StringComparison.Ordinal));
        }

        [Fact]
        public void Test_WordToHtml_HeadersFooters_RoundTripToExportedSectionIndex() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("First body");
            var second = doc.AddSection(SectionMarkValues.NextPage);
            second.AddParagraph("Second body");
            second.GetOrCreateHeader(HeaderFooterValues.Default).AddParagraph("Second header");

            string html = doc.ToHtml(new WordToHtmlOptions { ExportHeadersAndFooters = true });

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            Assert.True(roundTrip.Sections.Count >= 2);
            Assert.DoesNotContain(roundTrip.Sections[0].Header.Default?.Paragraphs ?? new List<WordParagraph>(),
                paragraph => string.Equals(paragraph.Text, "Second header", StringComparison.Ordinal));
            var header = Assert.IsAssignableFrom<WordHeader>(roundTrip.Sections[1].Header.Default);
            Assert.Contains(header.Paragraphs, paragraph => string.Equals(paragraph.Text, "Second header", StringComparison.Ordinal));
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
        public void Test_WordToHtml_ListDefinitions_ExportReusableCssWhenEnabled() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.HeadingIA1);
            list.AddItem("Intro");
            list.AddItem("Body");

            string defaultHtml = doc.ToHtml();

            Assert.DoesNotContain("data-word-list-definitions", defaultHtml, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("word-list-l0-ol-upper-roman", defaultHtml, StringComparison.OrdinalIgnoreCase);

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListDefinitions = true });

            Assert.Contains("data-word-list-definitions=\"true\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<ol start=\"1\" type=\"I\" class=\"word-list-l0-ol-upper-roman", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-word-list-level=\"0\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(".word-list-l0-ol-upper-roman", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list-style-type:upper-roman", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("list-style-position", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("--word-list-hanging", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("style=\"list-style-type:upper-roman\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_ListDefinitions_RoundTripUnderStrictCss() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.HeadingIA1);
            list.AddItem("Intro");
            list.AddItem("Body");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListDefinitions = true });
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);

            using var roundTrip = conversion.Value;

            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code.StartsWith("UnsupportedCss", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(roundTrip.Paragraphs, paragraph => string.Equals(paragraph.Text, "Intro", StringComparison.Ordinal));
        }

        [Fact]
        public void Test_WordToHtml_DefaultCss_RoundTripUnderStrictCss() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Normal paragraph");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeDefaultCss = true });
            var options = new HtmlToWordOptions {
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);

            using var roundTrip = conversion.Value;

            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code.StartsWith("UnsupportedCss", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(roundTrip.Paragraphs, paragraph => string.Equals(paragraph.Text, "Normal paragraph", StringComparison.Ordinal));
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
        public void Test_WordToHtml_InternationalListStyle() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.Numbered);
            list.Numbering.Levels[0]._level.NumberingFormat!.Val = NumberFormatValues.RussianLower;
            list.AddItem("alpha");
            list.AddItem("beta");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });

            Assert.Contains("<ol", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list-style-type:lower-russian", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_DashBulletListStyle() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.Bulleted);
            list.Numbering.Levels[0]._level.LevelText!.Val = "-";
            list.AddItem("One");
            list.AddItem("Two");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });

            Assert.Contains("<ul", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("type=\"disc\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list-style-type:'-'", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_EditorMarkerBulletListStyles() {
            using var doc = WordDocument.Create();
            var starList = doc.AddList(WordListStyle.Bulleted);
            starList.Numbering.Levels[0]._level.LevelText!.Val = "*";
            starList.AddItem("Star");
            var plusList = doc.AddList(WordListStyle.Bulleted);
            plusList.Numbering.Levels[0]._level.LevelText!.Val = "+";
            plusList.AddItem("Plus");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });

            Assert.Contains("list-style-type:'*'", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list-style-type:'+'", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_CustomBulletListStyle() {
            using var doc = WordDocument.Create();
            var checkList = doc.AddList(WordListStyle.Bulleted);
            checkList.Numbering.Levels[0]._level.LevelText!.Val = "✓";
            checkList.AddItem("Done");
            var diamondList = doc.AddList(WordListStyle.Bulleted);
            diamondList.Numbering.Levels[0]._level.LevelText!.Val = "◆";
            diamondList.AddItem("Diamond");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });

            Assert.Contains("list-style-type:'✓'", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list-style-type:'◆'", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("type=\"disc\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_CheckBoxControl() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("");
            paragraph.AddCheckBox(true, "Accept terms", "accept-terms");
            paragraph.AddText(" Accept terms");

            string html = doc.ToHtml();

            Assert.Contains("<input", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("type=\"checkbox\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("checked", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("aria-label=\"Accept terms\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-tag=\"accept-terms\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Accept terms", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("\u2611", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WordToHtml_TaskListCheckBoxControl() {
            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.Bulleted);
            var item = list.AddItem("", 0);
            item.AddCheckBox(false, "Open task", "open-task");
            item.AddText(" Open task");

            string html = doc.ToHtml();

            Assert.Contains("<ul", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<li><input", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("type=\"checkbox\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("disabled", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("checked", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Open task", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("\u2610", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WordToHtml_StructuredDocumentTag_ExportsTextInput() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Client: ");
            paragraph.AddStructuredDocumentTag("Contoso", "Client name", "client-name");

            string html = doc.ToHtml();

            Assert.Contains("<input", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("type=\"text\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("disabled", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("value=\"Contoso\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("aria-label=\"Client name\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-tag=\"client-name\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_MultilineStructuredDocumentTag_ExportsTextArea() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Notes: ");
            paragraph.AddStructuredDocumentTag("Line one\nLine two", "Review notes", "notes");

            string html = doc.ToHtml();

            Assert.Contains("<textarea", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("disabled", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("aria-label=\"Review notes\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-tag=\"notes\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Line one", html, StringComparison.Ordinal);
            Assert.Contains("Line two", html, StringComparison.Ordinal);
            Assert.DoesNotContain("value=\"Line one", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_MultilineStructuredDocumentTag_RoundTripsThroughTextArea() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Notes: ");
            paragraph.AddStructuredDocumentTag("Line one\nLine two", "Review notes", "notes");

            string html = doc.ToHtml();
            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            var control = Assert.Single(roundTrip.StructuredDocumentTags);
            Assert.Equal("Line one\nLine two", control.Text);
            Assert.Equal("Review notes", control.Alias);
            Assert.Equal("notes", control.Tag);
        }

        [Fact]
        public void Test_WordToHtml_DropDownList_ExportsSelect() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Priority: ");
            var dropDownList = paragraph.AddDropDownList(new[] { "Low", "Medium", "High" }, "Priority", "priority");
            dropDownList.SelectedValue = "High";

            string html = doc.ToHtml();

            Assert.Contains("<select", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("disabled", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("aria-label=\"Priority\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-tag=\"priority\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<option value=\"Low\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<option value=\"Medium\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<option value=\"High\" selected", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_ComboBox_ExportsInputWithDatalist() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Contact: ");
            paragraph.AddComboBox(new[] { "Email", "Phone" }, "Contact method", "contact-method", defaultValue: "Phone");

            string html = doc.ToHtml();

            Assert.Contains("type=\"text\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("list=\"word-combo-1\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("value=\"Phone\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("aria-label=\"Contact method\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-tag=\"contact-method\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<datalist id=\"word-combo-1\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<option value=\"Email\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<option value=\"Phone\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_DatePicker_ExportsDateInput() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Due: ");
            paragraph.AddDatePicker(new DateTime(2026, 7, 14), "Due date", "due-date");

            string html = doc.ToHtml();

            Assert.Contains("type=\"date\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("value=\"2026-07-14\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("aria-label=\"Due date\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-tag=\"due-date\"", html, StringComparison.OrdinalIgnoreCase);
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
        public void Test_WordToHtml_TableCellSpacing_ExportsBorderSpacingAndRoundTrips() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 1);
            table.StyleDetails!.CellSpacing = 240;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Spaced";
            table.Rows[0].Cells[0].Borders.LeftStyle = BorderValues.Single;
            table.Rows[0].Cells[0].Borders.RightStyle = BorderValues.Single;
            table.Rows[0].Cells[0].Borders.TopStyle = BorderValues.Single;
            table.Rows[0].Cells[0].Borders.BottomStyle = BorderValues.Single;

            string html = doc.ToHtml();

            Assert.Contains("border-spacing:12pt", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("border-collapse:separate", html, StringComparison.OrdinalIgnoreCase);

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            Assert.Equal((short)240, roundTrip.Tables[0].StyleDetails!.CellSpacing);
        }

        [Fact]
        public void Test_WordToHtml_TableCellVerticalAlignment_RoundTrips() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 1);
            table.Rows[0].Cells[0].VerticalAlignment = TableVerticalAlignmentValues.Center;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Centered";

            string html = doc.ToHtml();

            Assert.Contains("vertical-align:middle", html, StringComparison.OrdinalIgnoreCase);

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            Assert.Equal(TableVerticalAlignmentValues.Center, roundTrip.Tables[0].Rows[0].Cells[0].VerticalAlignment);
        }

        [Fact]
        public void Test_WordToHtml_TableCellCss() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 1);
            var cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
            cell.ShadingFillColorHex = "FF0000";
            cell.Borders.LeftStyle = BorderValues.Single;
            cell.Borders.RightStyle = BorderValues.Single;
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.BottomStyle = BorderValues.Single;
            cell.Borders.LeftColorHex = "00FF00";
            cell.Borders.RightColorHex = "00FF00";
            cell.Borders.TopColorHex = "00FF00";
            cell.Borders.BottomColorHex = "00FF00";
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
        public void Test_WordToHtml_CustomProperties_ExportAsTypedMetaWhenEnabled() {
            using var doc = WordDocument.Create();
            doc.CustomDocumentProperties["ReviewStatus"] = new WordCustomProperty("Approved");
            doc.CustomDocumentProperties["IsFinal"] = new WordCustomProperty(true);
            doc.CustomDocumentProperties["Score"] = new WordCustomProperty(6.15);
            doc.CustomDocumentProperties["ReviewedAt"] = new WordCustomProperty(new DateTime(2024, 1, 2, 3, 4, 5));

            string defaultHtml = doc.ToHtml();
            Assert.DoesNotContain("word:custom:ReviewStatus", defaultHtml, StringComparison.OrdinalIgnoreCase);

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeCustomProperties = true });

            Assert.Contains("name=\"word:custom:ReviewStatus\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("content=\"Approved\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-word-custom-property=\"ReviewStatus\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-property-type=\"Text\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("name=\"word:custom:IsFinal\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("content=\"true\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-property-type=\"YesNo\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("name=\"word:custom:Score\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("content=\"6.15\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-property-type=\"NumberDouble\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("name=\"word:custom:ReviewedAt\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("content=\"2024-01-02T03:04:05.0000000\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-property-type=\"DateTime\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_CustomProperties_RoundTripTypedMetaWhenEnabled() {
            using var doc = WordDocument.Create();
            doc.CustomDocumentProperties["ReviewStatus"] = new WordCustomProperty("Approved");
            doc.CustomDocumentProperties["IsFinal"] = new WordCustomProperty(true);
            doc.CustomDocumentProperties["Score"] = new WordCustomProperty(6.15);
            doc.CustomDocumentProperties["ReviewedAt"] = new WordCustomProperty(new DateTime(2024, 1, 2, 3, 4, 5));

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeCustomProperties = true });

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            Assert.Equal("Approved", roundTrip.CustomDocumentProperties["ReviewStatus"].Text);
            Assert.True(roundTrip.CustomDocumentProperties["IsFinal"].Bool);
            Assert.Equal(6.15, roundTrip.CustomDocumentProperties["Score"].NumberDouble);
            Assert.Equal(new DateTime(2024, 1, 2, 3, 4, 5), roundTrip.CustomDocumentProperties["ReviewedAt"].Date);
        }

        [Fact]
        public void Test_WordToHtml_BuiltinMetadata_RoundTrips() {
            using var doc = WordDocument.Create();
            doc.BuiltinDocumentProperties.Title = "Round Trip Title";
            doc.BuiltinDocumentProperties.Creator = "Ada";
            doc.BuiltinDocumentProperties.Description = "Round trip description";
            doc.BuiltinDocumentProperties.Keywords = "alpha,beta";
            doc.BuiltinDocumentProperties.Subject = "Metadata subject";
            doc.AddParagraph("Content");

            string html = doc.ToHtml();

            using var roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            Assert.Equal("Round Trip Title", roundTrip.BuiltinDocumentProperties.Title);
            Assert.Equal("Ada", roundTrip.BuiltinDocumentProperties.Creator);
            Assert.Equal("Round trip description", roundTrip.BuiltinDocumentProperties.Description);
            Assert.Equal("alpha,beta", roundTrip.BuiltinDocumentProperties.Keywords);
            Assert.Equal("Metadata subject", roundTrip.BuiltinDocumentProperties.Subject);
        }

        [Fact]
        public void Test_HtmlToWord_HeadStyles_AreParsedBeforeBodyImport() {
            string html = "<html><head><style>.title{font-weight:bold;font-size:24pt}</style></head><body><p class=\"title\">Title</p></body></html>";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            var paragraph = Assert.Single(doc.Paragraphs, paragraph => string.Equals(paragraph.Text, "Title", StringComparison.Ordinal));
            Assert.Equal(WordParagraphStyles.Heading2, paragraph.Style);
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
        public void Test_WordToHtml_StyleDefinitions_OmitAutomaticColors() {
            using var doc = WordDocument.Create();
            var style = new Style { Type = StyleValues.Paragraph, StyleId = "AutoColorStyle" };
            style.Append(new StyleName { Val = "Auto Color Style" });
            var runProperties = new StyleRunProperties();
            runProperties.Append(new Color { Val = "auto" });
            style.Append(runProperties);
            doc._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
            doc.AddParagraph("Automatic color").SetStyleId("AutoColorStyle");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeParagraphClasses = true });

            Assert.Contains(".AutoColorStyle {", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("color:#auto", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_WordToHtml_RunColorAndHighlightStyles() {
            using var doc = WordDocument.Create();
            var p = doc.AddParagraph();
            var run = p.AddText("Colored");
            run.ColorHex = "FF0000";
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

        [Fact]
        public void Test_WordToHtml_SectionMetadata_IsOptIn() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("First section");
            var second = doc.AddSection(SectionMarkValues.NextPage);
            second.PageOrientation = PageOrientationValues.Landscape;
            second.PageSettings.PageSize = WordPageSize.Letter;
            second.Margins.Top = 1440;
            second.Margins.Bottom = 720;
            second.Margins.Left = 1080;
            second.Margins.Right = 1080;
            second.AddParagraph("Second section");

            string defaultHtml = doc.ToHtml();
            Assert.DoesNotContain("class=\"word-section\"", defaultHtml, StringComparison.OrdinalIgnoreCase);

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeSectionMetadata = true });

            Assert.Contains("class=\"word-section\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-word-section=\"1\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-word-section=\"2\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-page-orientation=\"Landscape\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-page-size=\"Letter\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data-margin-top-twips=\"1440\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("padding:96px 72px 48px 72px", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("box-sizing", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("min-height", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("break-before:page", html, StringComparison.OrdinalIgnoreCase);
            Assert.True(html.IndexOf("First section", StringComparison.OrdinalIgnoreCase) < html.IndexOf("Second section", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_WordToHtml_SectionMetadata_RoundTripsWithoutExtraInitialSectionUnderStrictCss() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Only section");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeSectionMetadata = true });
            var options = new HtmlToWordOptions {
                SectionTagHandling = SectionTagHandling.WordSection,
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);

            using var roundTrip = conversion.Value;

            Assert.Single(roundTrip.Sections);
            Assert.Contains(roundTrip.Paragraphs, paragraph => string.Equals(paragraph.Text, "Only section", StringComparison.Ordinal));
            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code.StartsWith("UnsupportedCss", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void Test_WordToHtml_SectionMetadata_RoundTripsPageSetup() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("First section");
            var second = doc.AddSection(SectionMarkValues.NextPage);
            second.PageSettings.PageSize = WordPageSize.Letter;
            second.PageOrientation = PageOrientationValues.Landscape;
            second.Margins.Top = 1440;
            second.Margins.Bottom = 720;
            second.Margins.Left = 1080;
            second.Margins.Right = 1200;
            second.AddParagraph("Second section");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeSectionMetadata = true });
            var options = new HtmlToWordOptions {
                SectionTagHandling = SectionTagHandling.WordSection,
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Error
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);

            using var roundTrip = conversion.Value;

            Assert.Equal(2, roundTrip.Sections.Count);
            var roundTripSecond = roundTrip.Sections[1];
            Assert.Equal(PageOrientationValues.Landscape, roundTripSecond.PageOrientation);
            Assert.Equal(WordPageSize.Letter, roundTripSecond.PageSettings.PageSize);
            Assert.Equal(1440, roundTripSecond.Margins.Top);
            Assert.Equal(720, roundTripSecond.Margins.Bottom);
            Assert.Equal((UInt32Value)1080U, roundTripSecond.Margins.Left);
            Assert.Equal((UInt32Value)1200U, roundTripSecond.Margins.Right);
            var paragraph = Assert.Single(roundTripSecond.Paragraphs, paragraph => string.Equals(paragraph.Text, "Second section", StringComparison.Ordinal));
            Assert.Null(paragraph.IndentationBefore);
            Assert.Null(paragraph.IndentationAfter);
            Assert.Null(paragraph.LineSpacingBefore);
            Assert.Null(paragraph.LineSpacingAfter);
            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code.StartsWith("UnsupportedCss", StringComparison.OrdinalIgnoreCase));
        }
    }
}
