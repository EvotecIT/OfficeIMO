using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

            using var doc = OfficeIMO.Markdown.MarkdownDoc.Load(mdPath).ToWordDocument();
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
            Encoding latin1 = Encoding.GetEncoding("iso-8859-1");
            File.WriteAllText(mdPath, content, latin1);

            using var doc = OfficeIMO.Markdown.MarkdownDoc.Load(mdPath, encoding: latin1).ToWordDocument();
            string text = string.Join("\n", doc.Paragraphs.Select(p => p.Text));
            Assert.Contains("ol\u00e9", text);

            File.Delete(mdPath);
        }

        [Fact]
        public async Task Test_LoadFromMarkdownAsync_Path_DefaultEncoding() {
            string tempDir = Path.Combine(AppContext.BaseDirectory, "TempMarkdown");
            Directory.CreateDirectory(tempDir);
            string mdPath = Path.Combine(tempDir, "LoadFromPathDefaultAsync.md");
            string content = "# Title\r\n\r\nCaf\u00e9";
            File.WriteAllText(mdPath, content, Encoding.Unicode);

            using var doc = (await OfficeIMO.Markdown.MarkdownDoc.LoadAsync(mdPath)).ToWordDocument();
            string text = string.Join("\n", doc.Paragraphs.Select(p => p.Text));
            Assert.Contains("Caf\u00e9", text);

            File.Delete(mdPath);
        }
    }
}
