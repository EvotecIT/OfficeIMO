using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void WordToMarkdown_RendersUnderlineAndHighlight() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph();
            paragraph.AddText("underlined").Underline = UnderlineValues.Single;
            paragraph.AddText(" and ");
            paragraph.AddText("highlighted").Highlight = HighlightColorValues.Yellow;

            string markdownDefault = doc.ToMarkdown(new WordToMarkdownOptions());
            Assert.DoesNotContain("<u>", markdownDefault);
            Assert.DoesNotContain("==highlighted==", markdownDefault);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                EnableUnderline = true,
                EnableHighlight = true
            });

            Assert.Contains("<u>underlined</u>", markdown);
            Assert.Contains("==highlighted==", markdown);
        }
    }
}

