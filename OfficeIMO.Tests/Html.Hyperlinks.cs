using OfficeIMO.Word;

using OfficeIMO.Word.Html;
using System;
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
            Assert.Contains(doc.Bookmarks, b => string.Equals(b.Name, "_top", StringComparison.OrdinalIgnoreCase));



            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;


            Assert.NotNull(hyperlink);

            Assert.Equal("Back", hyperlink.Tooltip);

            Assert.Equal(TargetFrame._blank, hyperlink.TargetFrame);

            Assert.Equal("_top", hyperlink.Anchor);

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

        [Fact]
        public void Html_Hyperlinks_Normalizes_WwwLinks() {
            string html = "<p><a href=\"www.site.com\">Site</a></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;

            Assert.NotNull(hyperlink);
            Assert.Equal(new Uri("http://www.site.com/"), hyperlink!.Uri);
        }

        [Fact]
        public void Html_Hyperlinks_Normalizes_ProtocolRelativeLinks() {
            string html = "<p><a href=\"://www.site.com\">Site</a></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;

            Assert.NotNull(hyperlink);
            Assert.Equal(new Uri("http://www.site.com/"), hyperlink!.Uri);
        }

        [Fact]
        public void Html_Hyperlinks_InvalidHref_IsPlainText() {
            string html = "<p><a href=\"javascript:alert()\">Js</a></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Empty(doc.ParagraphsHyperLinks);
            Assert.Equal("Js", doc.Paragraphs[0].Text);
        }

        [Fact]
        public void Html_Hyperlinks_TopAnchor_CreatesBookmark() {
            string html = "<p>Start</p><p><a href=\"#top\">Move</a></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Contains(doc.Bookmarks, b => string.Equals(b.Name, "_top", StringComparison.OrdinalIgnoreCase));
            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;
            Assert.NotNull(hyperlink);
            Assert.Equal("_top", hyperlink!.Anchor);
        }

        [Fact]
        public void Html_Hyperlinks_NameAttribute_AddsBookmark() {
            string html = "<h1><a name=\"heading1\"></a>Heading</h1><p><a href=\"#heading1\">Go</a></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Contains(doc.Bookmarks, b => b.Name == "heading1");
            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;
            Assert.NotNull(hyperlink);
            Assert.Equal("heading1", hyperlink!.Anchor);
        }
    }

}

