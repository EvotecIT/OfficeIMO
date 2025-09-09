using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public async Task WordToHtmlConverter_Convert_EqualsAsync() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Test");
            var converter = new WordToHtmlConverter();
            string sync = converter.Convert(doc, new WordToHtmlOptions());
            string asyncResult = await converter.ConvertAsync(doc, new WordToHtmlOptions());
            Assert.Equal(sync, asyncResult);
        }

        [Fact]
        public async Task HtmlToWordConverter_Convert_EqualsAsync() {
            string html = "<p>Test</p>";
            var converter = new HtmlToWordConverter();
            using var syncDoc = converter.Convert(html, new HtmlToWordOptions());
            using var asyncDoc = await converter.ConvertAsync(html, new HtmlToWordOptions());
            Assert.Equal(syncDoc.Paragraphs.Count, asyncDoc.Paragraphs.Count);
            Assert.Equal(syncDoc.Paragraphs[0].Text, asyncDoc.Paragraphs[0].Text);
        }
    }
}
