using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointMarkdownTests {
        [Fact]
        public void CanApplyMarkdownToTextBoxRichTextRuns() {
            string filePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox textBox = slide.AddTextBox(string.Empty);

                var paragraphs = textBox.SetMarkdown("""
                    # Heading
                    - **Bold** and *italic* with `code`
                    3. Visit [OfficeIMO](https://example.com)
                    """);

                Assert.Equal(3, paragraphs.Count);
                Assert.True(paragraphs[0].Runs[0].Bold);
                Assert.Equal(28, paragraphs[0].Runs[0].FontSize);
                Assert.NotNull(paragraphs[1].Paragraph.ParagraphProperties?.GetFirstChild<CharacterBullet>());
                Assert.True(paragraphs[1].Runs[0].Bold);
                Assert.True(paragraphs[1].Runs[2].Italic);
                Assert.Equal("Consolas", paragraphs[1].Runs[4].FontName);
                Assert.NotNull(paragraphs[2].Paragraph.ParagraphProperties?.GetFirstChild<AutoNumberedBullet>());
                Assert.Equal(new Uri("https://example.com"), paragraphs[2].Runs.Last().Hyperlink);

                presentation.Save();
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
