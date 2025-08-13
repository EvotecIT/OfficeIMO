using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void AddHtmlToHeader_PlacesContentInHeader() {
            using WordDocument document = WordDocument.Create();
            document.AddHtmlToHeader("<p>Header content</p>");
            using var ms = new System.IO.MemoryStream();
            document.Save(ms);
            ms.Position = 0;
            using var docx = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
            var headerPart = docx.MainDocumentPart!.HeaderParts.First();
            Assert.Contains("Header content", headerPart.RootElement!.InnerText);
        }

        [Fact]
        public void AddHtmlToFooter_PlacesContentInFooter() {
            using WordDocument document = WordDocument.Create();
            document.AddHtmlToFooter("<p>Footer content</p>");
            using var ms = new System.IO.MemoryStream();
            document.Save(ms);
            ms.Position = 0;
            using var docx = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
            var footerPart = docx.MainDocumentPart!.FooterParts.First();
            Assert.Contains("Footer content", footerPart.RootElement!.InnerText);
        }
    }
}
