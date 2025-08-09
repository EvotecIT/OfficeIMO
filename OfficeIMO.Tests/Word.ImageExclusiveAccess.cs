using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ImageSaveToFileRequiresExclusiveAccess() {
            var docPath = Path.Combine(_directoryWithFiles, "ImageExclusive.docx");
            using var document = WordDocument.Create(docPath);
            document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"), 50, 50);

            var fileToSave = Path.Combine(_directoryWithFiles, "LockedImage.jpg");
            using (var lockStream = new FileStream(fileToSave, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite)) {
                Assert.Throws<IOException>(() => document.Images[0].SaveToFile(fileToSave));
            }

            document.Images[0].SaveToFile(fileToSave);
            Assert.True(new FileInfo(fileToSave).Length > 0);
        }
    }
}
