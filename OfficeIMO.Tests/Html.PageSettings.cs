using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_Respects_DefaultPageSettings() {
            string html = "<p>Hello</p>";
            var options = new HtmlToWordOptions {
                DefaultOrientation = PageOrientationValues.Landscape,
                DefaultPageSize = WordPageSize.A5
            };
            
            var doc = html.LoadFromHtml(options);
            
            Assert.Equal(PageOrientationValues.Landscape, doc.PageOrientation);
            Assert.Equal(WordPageSize.A5, doc.PageSettings.PageSize);
        }
    }
}