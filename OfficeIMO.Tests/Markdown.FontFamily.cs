using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void WordToMarkdown_UsesSpecifiedFontForCode() {
            using var doc = WordDocument.Create();
            string codeFont = FontResolver.Resolve("sans-serif")!;
            string normalFont = FontResolver.Resolve("serif")!;

            var paragraph = doc.AddParagraph();
            paragraph.AddText("code").SetFontFamily(codeFont);
            paragraph.AddText(" normal").SetFontFamily(normalFont);

            var options = new WordToMarkdownOptions {
                FontFamily = codeFont
            };

            string markdown = doc.ToMarkdown(options);

            Assert.Contains("`code`", markdown);
            Assert.Contains("normal", markdown);
            Assert.DoesNotContain("`normal`", markdown);
        }
    }
}