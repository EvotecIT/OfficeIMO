using System.IO;
using OfficeIMO.Converters;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public class ConverterExceptions {
        [Fact]
        public void HtmlToWordConverter_NullHtml_Throws() {
            using MemoryStream ms = new MemoryStream();
            Assert.Throws<ConversionException>(() => HtmlToWordConverter.Convert(null!, ms, new HtmlToWordOptions()));
        }

        [Fact]
        public void HtmlToWordConverter_NullOutput_Throws() {
            Assert.Throws<ConversionException>(() => HtmlToWordConverter.Convert("<p>test</p>", null!, new HtmlToWordOptions()));
        }

        [Fact]
        public void WordToHtmlConverter_NullInput_Throws() {
            Assert.Throws<ConversionException>(() => WordToHtmlConverter.Convert(null!));
        }

        [Fact]
        public void MarkdownToWordConverter_NullMarkdown_Throws() {
            using MemoryStream ms = new MemoryStream();
            Assert.Throws<ConversionException>(() => MarkdownToWordConverter.Convert(null!, ms, new MarkdownToWordOptions()));
        }

        [Fact]
        public void MarkdownToWordConverter_NullOutput_Throws() {
            Assert.Throws<ConversionException>(() => MarkdownToWordConverter.Convert("test", null!, new MarkdownToWordOptions()));
        }

        [Fact]
        public void WordToMarkdownConverter_NullInput_Throws() {
            Assert.Throws<ConversionException>(() => WordToMarkdownConverter.Convert(null!));
        }
    }
}

