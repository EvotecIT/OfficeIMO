using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_TextDecorationTags() {
            string html = "<p><s>strike</s><del>delete</del><ins>insert</ins><mark>mark</mark></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var runs = doc.Paragraphs;

            var strikeRun = runs.First(r => r.Text == "strike");
            Assert.True(strikeRun.Strike);

            var delRun = runs.First(r => r.Text == "delete");
            Assert.True(delRun.Strike);

            var insRun = runs.First(r => r.Text == "insert");
            Assert.Equal(UnderlineValues.Single, insRun.Underline);

            var markRun = runs.First(r => r.Text == "mark");
            Assert.Equal(HighlightColorValues.Yellow, markRun.Highlight);
        }
    }
}
