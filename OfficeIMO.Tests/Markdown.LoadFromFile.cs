using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void Test_LoadFromMarkdown_Path_DefaultEncoding() {
            string tempDir = Path.Combine(AppContext.BaseDirectory, "TempMarkdown");
            Directory.CreateDirectory(tempDir);
            string mdPath = Path.Combine(tempDir, "LoadFromPathDefault.md");
            string content = "# Title\r\n\r\nCaf\u00e9";
            File.WriteAllText(mdPath, content, Encoding.Unicode);

            using var doc = WordMarkdownConverterExtensions.LoadFromMarkdown(mdPath, encoding: null);
            string text = string.Join("\n", doc.Paragraphs.Select(p => p.Text));
            Assert.Contains("Caf\u00e9", text);

            File.Delete(mdPath);
        }

        [Fact]
        public void Test_LoadFromMarkdown_Path_CustomEncoding() {
            string tempDir = Path.Combine(AppContext.BaseDirectory, "TempMarkdown");
            Directory.CreateDirectory(tempDir);
            string mdPath = Path.Combine(tempDir, "LoadFromPathCustom.md");
            string content = "ol\u00e9";
            File.WriteAllText(mdPath, content, Encoding.Latin1);

            using var doc = WordMarkdownConverterExtensions.LoadFromMarkdown(mdPath, encoding: Encoding.Latin1);
            string text = string.Join("\n", doc.Paragraphs.Select(p => p.Text));
            Assert.Contains("ol\u00e9", text);

            File.Delete(mdPath);
        }
    }
}