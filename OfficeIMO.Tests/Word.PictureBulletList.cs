using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_PictureBulletList_FromFile() {
            var filePath = Path.Combine(_directoryWithFiles, "PictureBulletList.docx");
            using (var document = WordDocument.Create(filePath)) {
                var imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
                var list = document.AddPictureBulletList(imagePath);
                list.AddItem("Item1");
                list.AddItem("Item2");
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var numbering = document._wordprocessingDocument?.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                Assert.NotNull(numbering);
                Assert.NotNull(numbering!.Elements<NumberingPictureBullet>().FirstOrDefault());
                Assert.Single(document.Lists);
            }
        }

        [Fact]
        public void Test_PictureBulletList_FromStream() {
            var filePath = Path.Combine(_directoryWithFiles, "PictureBulletListStream.docx");
            using (var document = WordDocument.Create(filePath)) {
                var imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
                using var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var list = document.AddPictureBulletList(stream, "Kulek.jpg");
                list.AddItem("Item1");
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var numbering = document._wordprocessingDocument?.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                Assert.NotNull(numbering);
                Assert.NotNull(numbering!.Elements<NumberingPictureBullet>().FirstOrDefault());
                Assert.Single(document.Lists);
            }
        }
    }
}
