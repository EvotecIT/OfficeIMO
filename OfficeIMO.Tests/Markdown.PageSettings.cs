using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_Respects_DefaultPageSettings() {
            string md = "# Hello";
            var options = new MarkdownToWordOptions {
                DefaultOrientation = PageOrientationValues.Landscape,
                DefaultPageSize = WordPageSize.A5
            };
            
            var doc = md.LoadFromMarkdown(options);
            
            Assert.Equal(PageOrientationValues.Landscape, doc.PageOrientation);
            Assert.Equal(WordPageSize.A5, doc.PageSettings.PageSize);
        }
    }
}