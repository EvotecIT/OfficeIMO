using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_SaveAs_DoesNotChangeFilePath() {
            var originalPath = Path.Combine(_directoryWithFiles, "SaveAsOriginal.docx");
            var copyPath = Path.Combine(_directoryWithFiles, "SaveAsCopy.docx");

            File.Delete(originalPath);
            File.Delete(copyPath);

            using (var document = WordDocument.Create(originalPath)) {
                document.AddParagraph("Test");
                document.Save();

                Assert.Equal(originalPath, document.FilePath);

                using var newDoc = document.SaveAs(copyPath);
                Assert.Equal(originalPath, document.FilePath);
                Assert.Equal(copyPath, newDoc.FilePath);
                Assert.True(File.Exists(copyPath));
                Assert.Single(newDoc.Paragraphs);
            }

            using var loaded = WordDocument.Load(copyPath);
            Assert.Single(loaded.Paragraphs);
        }
    }
}
