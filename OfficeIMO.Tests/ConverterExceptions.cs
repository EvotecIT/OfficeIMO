using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public class ConverterExceptions {
        [Fact]
        public void LoadFromHtml_NullHtml_ThrowsArgumentNullException() {
            Assert.Throws<ArgumentNullException>(() => ((string)null!).LoadFromHtml());
        }

        [Fact]
        public void ToHtml_NullDocument_ThrowsArgumentNullException() {
            Assert.Throws<ArgumentNullException>(() => ((WordDocument)null!).ToHtml());
        }

        [Fact]
        public void LoadFromMarkdown_NullMarkdown_ThrowsArgumentNullException() {
            Assert.Throws<ArgumentNullException>(() => ((string)null!).LoadFromMarkdown());
        }

        [Fact]
        public void ToMarkdown_NullDocument_ThrowsArgumentNullException() {
            Assert.Throws<ArgumentNullException>(() => ((WordDocument)null!).ToMarkdown());
        }
    }
}

