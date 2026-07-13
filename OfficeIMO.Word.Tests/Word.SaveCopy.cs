using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveCopy_DoesNotChangeAssociatedPath() {
            var originalPath = Path.Combine(_directoryWithFiles, "SaveCopyOriginal.docx");
            var copyPath = Path.Combine(_directoryWithFiles, "SaveCopy.docx");

            File.Delete(originalPath);
            File.Delete(copyPath);

            using (var document = WordDocument.Create(originalPath)) {
                document.AddParagraph("Test");
                document.Save();

                Assert.Equal(originalPath, document.FilePath);

                document.SaveCopy(copyPath);
                Assert.Equal(originalPath, document.FilePath);
                Assert.True(File.Exists(copyPath));
                using var newDoc = WordDocument.Load(copyPath);
                Assert.Equal(copyPath, newDoc.FilePath);
                Assert.Single(newDoc.Paragraphs);
            }

            using var loaded = WordDocument.Load(copyPath);
            Assert.Single(loaded.Paragraphs);
        }

        [Fact]
        public async Task SaveCopyAsync_DoesNotChangeAssociatedPath() {
            string originalPath = Path.Combine(_directoryWithFiles, "SaveCopyAsyncOriginal.docx");
            string copyPath = Path.Combine(_directoryWithFiles, "SaveCopyAsyncCopy.docx");
            File.Delete(originalPath);
            File.Delete(copyPath);

            await using WordDocument document = WordDocument.Create(originalPath);
            document.AddParagraph("Async copy");
            document.Save();
            await document.SaveCopyAsync(copyPath);

            Assert.Equal(originalPath, document.FilePath);
            using WordDocument copy = WordDocument.Load(copyPath);
            Assert.Equal("Async copy", Assert.Single(copy.Paragraphs).Text);
        }
    }
}
