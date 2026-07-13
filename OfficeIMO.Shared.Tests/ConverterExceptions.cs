using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public class ConverterExceptions {
        [Fact]
        public void ToWordDocument_NullHtml_ThrowsArgumentNullException() {
            Assert.Throws<ArgumentNullException>(() => OfficeIMO.Html.HtmlConversionDocument.Parse((string)null!));
        }

        [Fact]
        public void ToHtml_NullDocument_ThrowsArgumentNullException() {
            Assert.Throws<ArgumentNullException>(() => ((WordDocument)null!).ToHtml());
        }

        [Fact]
        public void LoadFromMarkdown_NullMarkdown_ThrowsArgumentNullException() {
            Assert.Throws<ArgumentNullException>(() => OfficeIMO.Markdown.MarkdownReader.Parse((string)null!));
        }

        [Fact]
        public void ToMarkdown_NullDocument_ThrowsArgumentNullException() {
            Assert.Throws<ArgumentNullException>(() => ((WordDocument)null!).ToMarkdown());
        }
    }
}
