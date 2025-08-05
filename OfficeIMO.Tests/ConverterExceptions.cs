using System.IO;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public class ConverterExceptions {
        [Fact]
        public void HtmlToWordConverter_NullHtml_Throws() {
            // TODO: This test now passes after adding null check
            using MemoryStream ms = new MemoryStream();
            Assert.Throws<ArgumentNullException>(() => ((string)null!).LoadFromHtml());
        }

        [Fact]
        public void HtmlToWordConverter_NullOutput_Throws() {
            // This test is no longer valid as LoadFromHtml returns WordDocument
        }

        [Fact]
        public void WordToHtmlConverter_NullInput_Throws() {
            // TODO: This test now passes after adding null check
            Assert.Throws<ArgumentNullException>(() => ((WordDocument)null!).ToHtml());
        }

        [Fact]
        public void MarkdownToWordConverter_NullMarkdown_Throws() {
            // TODO: This test now passes after adding null check
            using MemoryStream ms = new MemoryStream();
            Assert.Throws<ArgumentNullException>(() => ((string)null!).LoadFromMarkdown());
        }

        [Fact]
        public void MarkdownToWordConverter_NullOutput_Throws() {
            // This test is no longer valid as LoadFromMarkdown returns WordDocument
        }

        [Fact]
        public void WordToMarkdownConverter_NullInput_Throws() {
            // TODO: This test now passes after adding null check
            Assert.Throws<ArgumentNullException>(() => ((WordDocument)null!).ToMarkdown());
        }
    }
}

