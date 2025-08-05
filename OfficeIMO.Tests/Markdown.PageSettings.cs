using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_Respects_DefaultPageSettings() {
            string md = "Hello";
            using MemoryStream ms = new MemoryStream();
            var options = new MarkdownToWordOptions {
                DefaultOrientation = PageOrientationValues.Landscape,
                DefaultPageSize = WordPageSize.A5
            };
            MarkdownToWordConverter.Convert(md, ms, options);

            ms.Position = 0;
            using WordDocument doc = WordDocument.Load(ms);
            Assert.Equal(PageOrientationValues.Landscape, doc.PageOrientation);
            Assert.Equal(WordPageSize.A5, doc.PageSettings.PageSize);
        }
    }
}
