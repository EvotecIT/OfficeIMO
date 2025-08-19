using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public class ImageEmbedderTests {
        [Fact]
        public void Test_ImageEmbedder_AddsImage() {
            using MemoryStream ms = new MemoryStream();
            using WordprocessingDocument doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true);
            MainDocumentPart mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            Run run = ImageEmbedder.CreateImageRun(mainPart, assetPath);
            Assert.NotNull(mainPart.Document);
            Assert.NotNull(mainPart.Document.Body);
            mainPart.Document.Body!.Append(new Paragraph(run));
            mainPart.Document.Save();

            Assert.NotEmpty(mainPart.ImageParts);
        }
    }
}
