using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_Respects_DefaultPageSettings() {
            string html = "<p>Hello</p>";
            using MemoryStream ms = new MemoryStream();
            var options = new HtmlToWordOptions {
                DefaultOrientation = PageOrientationValues.Landscape,
                DefaultPageSize = WordPageSize.A5
            };
            HtmlToWordConverter.Convert(html, ms, options);

            ms.Position = 0;
            using WordDocument doc = WordDocument.Load(ms);
            Assert.Equal(PageOrientationValues.Landscape, doc.PageOrientation);
            Assert.Equal(WordPageSize.A5, doc.PageSettings.PageSize);
        }
    }
}
