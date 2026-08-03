using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Linq;
using Xunit;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

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

        [Fact]
        public void Test_ImageEmbedder_ReservesUniqueIdsBeforeDetachedRunsAreAppended() {
            using MemoryStream ms = new MemoryStream();
            using WordprocessingDocument doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true);
            MainDocumentPart mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");

            Run first = ImageEmbedder.CreateImageRun(mainPart, assetPath);
            Run second = ImageEmbedder.CreateImageRun(mainPart, assetPath);
            mainPart.Document.Body!.Append(new Paragraph(first), new Paragraph(second));

            uint[] ids = mainPart.Document.Descendants<DW.DocProperties>().Select(properties => properties.Id!.Value)
                .Concat(mainPart.Document.Descendants<PIC.NonVisualDrawingProperties>().Select(properties => properties.Id!.Value))
                .ToArray();
            Assert.Equal(4, ids.Length);
            Assert.Equal(ids.Length, ids.Distinct().Count());
            Assert.All(ids, id => Assert.True(id > 0U));
        }
    }
}
