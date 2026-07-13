using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void WordToHtml_UsesDocumentConversionSurface() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Test");
            string html = doc.ToHtml(new WordToHtmlOptions());
            Assert.Contains("Test", html, StringComparison.Ordinal);
        }

        [Fact]
        public async Task HtmlToWordConverter_Convert_EqualsAsync() {
            string html = "<p>Test</p>";
            OfficeIMO.Html.HtmlConversionDocument source = OfficeIMO.Html.HtmlConversionDocument.Parse(html);
            using var syncDoc = source.ToWordDocument(new HtmlToWordOptions());
            using var asyncDoc = await source.ToWordDocumentAsync(new HtmlToWordOptions());
            Assert.Equal(syncDoc.Paragraphs.Count, asyncDoc.Paragraphs.Count);
            Assert.Equal(syncDoc.Paragraphs[0].Text, asyncDoc.Paragraphs[0].Text);
        }
    }
}
