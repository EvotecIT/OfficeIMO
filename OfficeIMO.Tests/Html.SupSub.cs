using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_SupSub_RoundTrip() {
            string html = "<p>H<sub>2</sub>O</p><p>Note<sup>1</sup></p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var docRuns = doc.Paragraphs;

            var subRun = docRuns.First(r => r.Text == "2");
            Assert.Equal(VerticalPositionValues.Subscript, subRun.VerticalTextAlignment);

            var supRun = docRuns.First(r => r.Text == "1");
            Assert.Equal(VerticalPositionValues.Superscript, supRun.VerticalTextAlignment);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<sub>2</sub>", roundTrip);
            Assert.Contains("<sup>1</sup>", roundTrip);
        }
    }
}

