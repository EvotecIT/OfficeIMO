using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void Header_AddHtml_PlacesContentInHeader() {
            using WordDocument document = WordDocument.Create();
            document.Sections.Last().PageSetup.Header.AddHtml(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>Header content</p>"));
            using var ms = new System.IO.MemoryStream();
            document.Save(ms);
            ms.Position = 0;
            using var docx = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
            var headerPart = docx.MainDocumentPart!.HeaderParts.First();
            Assert.Contains("Header content", headerPart.RootElement!.InnerText);
        }

        [Fact]
        public void Header_AddHtml_InlineQuoteStaysInHeader() {
            using WordDocument document = WordDocument.Create();
            document.Sections.Last().PageSetup.Header.AddHtml(OfficeIMO.Html.HtmlConversionDocument.Parse("<q>Quoted header</q>"));

            string innerText = GetHeaderInnerText(document, WordHeaderFooterType.Default);

            Assert.Contains("Quoted header", innerText);
            Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text.Contains("Quoted header", StringComparison.Ordinal));
        }

        [Fact]
        public void Footer_AddHtml_PlacesContentInFooter() {
            using WordDocument document = WordDocument.Create();
            document.Sections.Last().PageSetup.Footer.AddHtml(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>Footer content</p>"));
            using var ms = new System.IO.MemoryStream();
            document.Save(ms);
            ms.Position = 0;
            using var docx = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
            var footerPart = docx.MainDocumentPart!.FooterParts.First();
            Assert.Contains("Footer content", footerPart.RootElement!.InnerText);
        }

        [Fact]
        public void Footer_AddHtml_InsertsIntoFirstFooter() {
            AssertFooterCreated(WordHeaderFooterType.First, "Sync first footer", doc => doc.PageSetup.DifferentFirstPage = true);
        }

        [Fact]
        public void Footer_AddHtml_InsertsIntoEvenFooter() {
            AssertFooterCreated(WordHeaderFooterType.Even, "Sync even footer", doc => doc.PageSetup.DifferentOddAndEvenPages = true);
        }

        [Fact]
        public void Header_AddHtml_InsertsIntoFirstHeader() {
            AssertHeaderCreated(WordHeaderFooterType.First, "Sync first header", doc => doc.PageSetup.DifferentFirstPage = true);
        }

        [Fact]
        public void Header_AddHtml_InsertsIntoEvenHeader() {
            AssertHeaderCreated(WordHeaderFooterType.Even, "Sync even header", doc => doc.PageSetup.DifferentOddAndEvenPages = true);
        }

        private static void AssertFooterCreated(WordHeaderFooterType footerType, string expectedText, Action<WordDocument> configure) {
            using WordDocument document = WordDocument.Create();
            configure(document);

            string html = $"<p>{expectedText}</p>";
            document.Sections.Last().PageSetup.GetOrCreateFooter(footerType).AddHtml(OfficeIMO.Html.HtmlConversionDocument.Parse(html));

            Assert.NotNull(document.Sections.Last().PageSetup.GetFooter(footerType));

            string innerText = GetFooterInnerText(document, footerType);
            Assert.Contains(expectedText, innerText);
        }

        private static void AssertHeaderCreated(WordHeaderFooterType headerType, string expectedText, Action<WordDocument> configure) {
            using WordDocument document = WordDocument.Create();
            configure(document);

            string html = $"<p>{expectedText}</p>";
            document.Sections.Last().PageSetup.GetOrCreateHeader(headerType).AddHtml(OfficeIMO.Html.HtmlConversionDocument.Parse(html));

            Assert.NotNull(document.Sections.Last().PageSetup.GetHeader(headerType));

            string innerText = GetHeaderInnerText(document, headerType);
            Assert.Contains(expectedText, innerText);
        }

    }
}
