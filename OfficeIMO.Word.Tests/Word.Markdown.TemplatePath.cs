using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordMarkdownTemplatePathTests {
    [Fact]
    public void TemplatePathCreatesADetachedOutputAndKeepsTheSourceUnchanged() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-markdown-template-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "template.docx");
            using (WordDocument template = WordDocument.Create()) {
                template.AddParagraph("Retained template paragraph");
                template.Save(path);
            }
            byte[] original = File.ReadAllBytes(path);
            using MarkdownDocument markdown = MarkdownDocument.Parse("Inserted **content**");
            var options = new MarkdownWordExportOptions { TemplatePath = path };
            using WordDocument output = markdown.ToWord(options, out _);
            Assert.True(string.IsNullOrEmpty(output.FilePath));
            Assert.Contains(output.Paragraphs, paragraph => paragraph.Text.Contains("Retained template paragraph"));
            Assert.Contains("Inserted content", string.Concat(output.Paragraphs.Select(paragraph => paragraph.Text)));
            Assert.Equal(original, File.ReadAllBytes(path));
            using (FileStream source = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                Assert.Equal(original.Length, source.Length);
            }
            output.AddParagraph("Caller-owned output");
            Assert.NotEmpty(output.ToBytes());
            Assert.Equal(path, options.TemplatePath);
            Assert.Null(options.Template);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally {
            Directory.Delete(root, true);
        }
    }
}
