using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void AddHtmlToBody_InsertsAtBeginning() {
            using WordDocument document = WordDocument.Create();
            document.AddHtmlToBody("<p>First</p>");
            document.AddParagraph("Second");

            Assert.Equal("First", document.Paragraphs[0].Text);
            Assert.Equal("Second", document.Paragraphs[1].Text);
        }

        [Fact]
        public void AddHtmlToBody_InsertsInMiddle() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Start");
            document.AddHtmlToBody("<p>Middle</p>");
            document.AddParagraph("End");

            Assert.Equal("Start", document.Paragraphs[0].Text);
            Assert.Equal("Middle", document.Paragraphs[1].Text);
            Assert.Equal("End", document.Paragraphs[2].Text);
        }

        [Fact]
        public void AddHtmlToBody_InsertsAtEnd() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Start");
            document.AddParagraph("Middle");
            document.AddHtmlToBody("<p>End</p>");

            Assert.Equal("Start", document.Paragraphs[0].Text);
            Assert.Equal("Middle", document.Paragraphs[1].Text);
            Assert.Equal("End", document.Paragraphs[2].Text);
        }
    }
}

