using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void SaveAsMarkdown_WithoutOptions_WritesImagesNextToTheFile() {
            string directory = Path.Combine(Path.GetTempPath(), "officeimo-markdown-saved-images-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try {
                using var word = WordDocument.Create();
                byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
                using (var source = new MemoryStream(png)) word.AddParagraph().AddImage(source, "chart.png", width: OfficeLength.Pixels(32), height: OfficeLength.Pixels(32));
                string output = Path.Combine(directory, "report.md");

                word.SaveAsMarkdown(output);

                string markdown = File.ReadAllText(output);
                Assert.DoesNotContain("data:image/", markdown, StringComparison.Ordinal);
                Assert.Contains("](chart", markdown, StringComparison.Ordinal);
                Assert.NotEmpty(Directory.GetFiles(directory, "chart*.png"));
                // In-memory conversion has no folder to write to and keeps data URIs.
                Assert.Contains("data:image/png;base64", word.ToMarkdown(new WordMarkdownExportOptions()).ToMarkdown(), StringComparison.Ordinal);
            } finally {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}
