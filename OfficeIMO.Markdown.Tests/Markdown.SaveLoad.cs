using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void Test_SaveAsMarkdown_FileAndRead() {
            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            try {
                string filePath = Path.Combine(tempDir, "FileSave.md");
                using var doc = WordDocument.Create();
                doc.AddParagraph("File save");
                doc.SaveAsMarkdown(filePath);

                Assert.True(File.Exists(filePath));
                string content = File.ReadAllText(filePath, Encoding.UTF8);
                Assert.Contains("File save", content);
            } finally {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public void Test_SaveAsMarkdown_StreamAndRead() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Stream save");

            using var stream = new MemoryStream();
            doc.SaveAsMarkdown(stream);
            stream.Position = 0;
            using var reader = new StreamReader(stream, Encoding.UTF8);
            string content = reader.ReadToEnd();
            Assert.Contains("Stream save", content);
        }

        [Fact]
        public async Task Test_SaveLoad_RoundTrip_FromFilePath() {
            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            try {
                string filePath = Path.Combine(tempDir, "RoundTrip.md");
                using var originalDoc = WordDocument.Create();
                originalDoc.AddParagraph("Roundtrip content");
                originalDoc.SaveAsMarkdown(filePath);

                using var loaded = await filePath.LoadFromMarkdownAsync();
                Assert.NotEmpty(loaded.Paragraphs);
                Assert.Contains("Roundtrip content", loaded.Paragraphs.Select(p => p.Text));
            } finally {
                Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task Test_LoadFromMarkdown_Stream_CustomEncoding() {
            string text = "Zażółć gęślą jaźń";
            using var stream = new MemoryStream();
            using (var writer = new StreamWriter(stream, Encoding.Unicode, 1024, leaveOpen: true)) {
                writer.Write(text);
                writer.Flush();
            }
            stream.Position = 0;
            using var doc = await stream.LoadFromMarkdownAsync();
            Assert.Contains(text, doc.Paragraphs.Select(p => p.Text));
        }
    }
}
