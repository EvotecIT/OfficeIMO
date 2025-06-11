using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ImageLocation_Document() {
            var filePath = Path.Combine(_directoryWithFiles, "ImageLocationDocument.docx");
            using var document = WordDocument.Create(filePath);

            var paragraph = document.AddParagraph();
            paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);

            var mainPart = document._wordprocessingDocument.MainDocumentPart!;

            Assert.Single(mainPart.ImageParts);
            Assert.Empty(mainPart.HeaderParts.SelectMany(h => h.ImageParts));
            Assert.Empty(mainPart.FooterParts.SelectMany(f => f.ImageParts));

            document.Save(false);
        }

        [Fact]
        public void Test_ImageLocation_Header() {
            var filePath = Path.Combine(_directoryWithFiles, "ImageLocationHeader.docx");
            using var document = WordDocument.Create(filePath);
            document.AddHeadersAndFooters();

            var paragraph = document.Header.Default.AddParagraph();
            paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);

            var mainPart = document._wordprocessingDocument.MainDocumentPart!;
            var headerPart = mainPart.HeaderParts.First();

            Assert.Single(headerPart.ImageParts);
            Assert.Empty(mainPart.ImageParts);
            Assert.Empty(mainPart.FooterParts.First().ImageParts);

            document.Save(false);
        }

        [Fact]
        public void Test_ImageLocation_Footer() {
            var filePath = Path.Combine(_directoryWithFiles, "ImageLocationFooter.docx");
            using var document = WordDocument.Create(filePath);
            document.AddHeadersAndFooters();

            var paragraph = document.Footer.Default.AddParagraph();
            paragraph.AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);

            var mainPart = document._wordprocessingDocument.MainDocumentPart!;
            var footerPart = mainPart.FooterParts.First();

            Assert.Single(footerPart.ImageParts);
            Assert.Empty(mainPart.ImageParts);
            Assert.Empty(mainPart.HeaderParts.First().ImageParts);

            document.Save(false);
        }
    }
}
