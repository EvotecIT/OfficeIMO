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

        [Fact]
        public void HtmlToWord_BlockStyles_MapPageBreakBeforeAndAfter() {
            string html = "<p>Intro</p><p style=\"break-before: page\">Next</p><p style=\"page-break-after: always\">Tail</p><p>After</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.False(doc.Paragraphs[0].PageBreakBefore);
            Assert.True(doc.Paragraphs[1].PageBreakBefore);
            Assert.Null(doc.Paragraphs[2].PageBreak);
            Assert.NotNull(doc.Paragraphs[3].PageBreak);
            Assert.Equal(BreakValues.Page, doc.Paragraphs[3].PageBreak!.BreakType);
            Assert.Null(doc.Paragraphs[4].PageBreak);
        }

        [Fact]
        public void HtmlToWord_ContainerBreakAfter_AppliesToLastGeneratedParagraphOnly() {
            string html = "<div style=\"break-after: page\"><p>First</p><p>Second</p></div><p>Third</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Null(doc.Paragraphs[0].PageBreak);
            Assert.Null(doc.Paragraphs[1].PageBreak);
            Assert.NotNull(doc.Paragraphs[2].PageBreak);
            Assert.Equal(BreakValues.Page, doc.Paragraphs[2].PageBreak!.BreakType);
            Assert.Null(doc.Paragraphs[3].PageBreak);
        }
    }
}
