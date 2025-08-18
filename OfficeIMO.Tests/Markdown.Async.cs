using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public async Task Test_SaveAsMarkdownAsync_FileAndLoadAsync() {
            string tempDir = Path.Combine(AppContext.BaseDirectory, "TempMarkdown");
            Directory.CreateDirectory(tempDir);
            string mdPath = Path.Combine(tempDir, "AsyncFile.md");
            if (File.Exists(mdPath)) File.Delete(mdPath);

            await using (var doc = WordDocument.Create()) {
                doc.AddParagraph("Async file");
                await doc.SaveAsMarkdownAsync(mdPath);
            }

            Assert.True(File.Exists(mdPath));

            await using (var doc = await mdPath.LoadFromMarkdownAsync()) {
                Assert.True(doc.Paragraphs.Count >= 1);
                Assert.Contains("Async file", string.Join("\n", doc.Paragraphs.Select(p => p.Text)));
            }

            File.Delete(mdPath);
        }

        [Fact]
        public async Task Test_SaveAsMarkdownAsync_StreamAndLoadAsync() {
            await using var doc = WordDocument.Create();
            doc.AddParagraph("Async stream");

            await using var stream = new MemoryStream();
            await doc.SaveAsMarkdownAsync(stream);
            Assert.True(stream.CanRead);
            stream.Position = 0;

            await using var loaded = await stream.LoadFromMarkdownAsync();
            Assert.True(loaded.Paragraphs.Count >= 1);
            Assert.Contains("Async stream", string.Join("\n", loaded.Paragraphs.Select(p => p.Text)));

            Assert.True(stream.CanRead);
        }
    }
}

