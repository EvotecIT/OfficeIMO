using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void WordToMarkdown_EscapesSpecialCharacters() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Characters: * _ [ ] ( ) # + - . ! \\ >");

            string md = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.Contains("\\*", md);
            Assert.Contains("\\_", md);
            Assert.Contains("\\[", md);
            Assert.Contains("\\]", md);
            Assert.Contains("\\(", md);
            Assert.Contains("\\)", md);
            Assert.Contains("\\#", md);
            Assert.Contains("\\+", md);
            Assert.Contains("\\-", md);
            Assert.Contains("\\.", md);
            Assert.Contains("\\!", md);
            Assert.Contains("\\\\", md);
            Assert.Contains("\\>", md);
        }
    }
}

