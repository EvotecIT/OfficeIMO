using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingPictureControl() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithPictureControl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
                var pc = document.AddParagraph().AddPictureControl(imagePath, 100, 100, "PC", "PCTag");

                Assert.Single(document.PictureControls);
                Assert.Equal("PC", pc.Alias);
                Assert.Equal("PCTag", pc.Tag);

                document.Save(false);
                Assert.False(HasUnexpectedElements(document), "Document has unexpected elements. Order of elements matters!");
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.PictureControls);
                var pic = document.GetPictureControlByAlias("PC");
                Assert.NotNull(pic);
                Assert.Equal("PCTag", document.GetPictureControlByTag("PCTag")?.Tag);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.PictureControls[0].Remove();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.PictureControls);
            }
        }
    }
}
