using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_FontTagAttributes() {
            string html = "<p><font color=\"#00FF00\" size=\"5\">Green</font></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();

            Assert.Equal("00ff00", run.ColorHex);
            Assert.Equal(18, run.FontSize);
        }
    }
}
