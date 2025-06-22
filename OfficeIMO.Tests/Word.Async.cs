using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public async Task Test_WordSaveLoadAsync() {
            var filePath = Path.Combine(_directoryWithFiles, "AsyncWord.docx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = WordDocument.Create(filePath)) {
                document.AddParagraph("Async");
                await document.SaveAsync();
            }

            Assert.True(File.Exists(filePath));

            using (var document = await WordDocument.LoadAsync(filePath)) {
                Assert.Single(document.Paragraphs);
            }

            File.Delete(filePath);
        }
    }
}
