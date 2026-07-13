using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void AddHtmlToHeader_PlacesContentInHeader() {
            using WordDocument document = WordDocument.Create();
            document.AddHtmlToHeader(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>Header content</p>"));
            using var ms = new System.IO.MemoryStream();
            document.Save(ms);
            ms.Position = 0;
            using var docx = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
            var headerPart = docx.MainDocumentPart!.HeaderParts.First();
            Assert.Contains("Header content", headerPart.RootElement!.InnerText);
        }

        [Fact]
        public void AddHtmlToHeader_InlineQuoteStaysInHeader() {
            using WordDocument document = WordDocument.Create();
            document.AddHtmlToHeader(OfficeIMO.Html.HtmlConversionDocument.Parse("<q>Quoted header</q>"));

            string innerText = GetHeaderInnerText(document, HeaderFooterValues.Default);

            Assert.Contains("Quoted header", innerText);
            Assert.DoesNotContain(document.Paragraphs, paragraph => paragraph.Text.Contains("Quoted header", StringComparison.Ordinal));
        }

        [Fact]
        public void AddHtmlToFooter_PlacesContentInFooter() {
            using WordDocument document = WordDocument.Create();
            document.AddHtmlToFooter(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>Footer content</p>"));
            using var ms = new System.IO.MemoryStream();
            document.Save(ms);
            ms.Position = 0;
            using var docx = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
            var footerPart = docx.MainDocumentPart!.FooterParts.First();
            Assert.Contains("Footer content", footerPart.RootElement!.InnerText);
        }

        [Fact]
        public void AddHtmlToFooter_CreatesFirstFooter() {
            AssertFooterCreated(HeaderFooterValues.First, "Sync first footer", doc => doc.DifferentFirstPage = true);
        }

        [Fact]
        public void AddHtmlToFooter_CreatesEvenFooter() {
            AssertFooterCreated(HeaderFooterValues.Even, "Sync even footer", doc => doc.DifferentOddAndEvenPages = true);
        }

        [Fact]
        public void AddHtmlToHeader_CreatesFirstHeader() {
            AssertHeaderCreated(HeaderFooterValues.First, "Sync first header", doc => doc.DifferentFirstPage = true);
        }

        [Fact]
        public void AddHtmlToHeader_CreatesEvenHeader() {
            AssertHeaderCreated(HeaderFooterValues.Even, "Sync even header", doc => doc.DifferentOddAndEvenPages = true);
        }

        private static void AssertFooterCreated(HeaderFooterValues footerType, string expectedText, Action<WordDocument> configure) {
            using WordDocument document = WordDocument.Create();
            configure(document);

            string html = $"<p>{expectedText}</p>";
            document.AddHtmlToFooter(OfficeIMO.Html.HtmlConversionDocument.Parse(html), footerType);

            var section = document.Sections.Last();
            var footers = section.Footer;
            Assert.NotNull(footers);
            Assert.NotNull(ResolveFooter(footers!, footerType));

            string innerText = GetFooterInnerText(document, footerType);
            Assert.Contains(expectedText, innerText);
        }

        private static void AssertHeaderCreated(HeaderFooterValues headerType, string expectedText, Action<WordDocument> configure) {
            using WordDocument document = WordDocument.Create();
            configure(document);

            string html = $"<p>{expectedText}</p>";
            document.AddHtmlToHeader(OfficeIMO.Html.HtmlConversionDocument.Parse(html), headerType);

            var section = document.Sections.Last();
            var headers = section.Header;
            Assert.NotNull(headers);
            Assert.NotNull(ResolveHeader(headers!, headerType));

            string innerText = GetHeaderInnerText(document, headerType);
            Assert.Contains(expectedText, innerText);
        }

    }
}
