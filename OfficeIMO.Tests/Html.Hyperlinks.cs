using OfficeIMO.Word;

using OfficeIMO.Word.Html;
using System.Linq;

using Xunit;



namespace OfficeIMO.Tests {

    public partial class Html {

        [Fact]

        public void Html_Hyperlinks_Title_And_Target_External() {

            string html = "<p><a href=\"https://example.com\" title=\"Example\" target=\"_self\">Example</a></p>";



            var doc = html.LoadFromHtml(new HtmlToWordOptions());



            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;


            Assert.NotNull(hyperlink);

            Assert.Equal("Example", hyperlink.Tooltip);

            Assert.Equal(TargetFrame._self, hyperlink.TargetFrame);

        }



        [Fact]

        public void Html_Hyperlinks_InternalAnchor() {

            string html = "<p id=\"top\">Top</p><p><a href=\"#top\" title=\"Back\" target=\"_blank\">Back</a></p>";



            var doc = html.LoadFromHtml(new HtmlToWordOptions());



            Assert.Contains(doc.Bookmarks, b => b.Name == "top");



            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;


            Assert.NotNull(hyperlink);

            Assert.Equal("Back", hyperlink.Tooltip);

            Assert.Equal(TargetFrame._blank, hyperlink.TargetFrame);

            Assert.Equal("top", hyperlink.Anchor);

        }


        [Fact]
        public void Html_Hyperlinks_PreserveInlineFormatting() {
            string html = "<p><a href=\"https://example.com\"><strong>Go</strong> now</a></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var paragraph = doc.ParagraphsHyperLinks[0];
            var runs = paragraph.GetRuns().Where(r => !r.IsBreak).ToList();
            var text = string.Concat(runs.Select(r => r.Text));

            Assert.Equal("Go now", text);
            Assert.Contains(runs, r => r.Text == "Go" && r.Bold);
        }
    }

}

