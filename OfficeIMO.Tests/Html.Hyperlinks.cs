using OfficeIMO.Word;
using OfficeIMO.Word.Html;
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


            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Contains(doc.Bookmarks, b => b.Name == "intro");

            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;

            Assert.NotNull(hyperlink);
            Assert.Equal("intro", hyperlink.Anchor);
        }

        [Fact]
        public void Html_Hyperlinks_LinkToTopUsesTopBookmark() {
            string html = "<p>Content</p><p><a href=\"#top\">Top</a></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Contains(doc.Bookmarks, b => b.Name == "_top");

            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;

            Assert.NotNull(hyperlink);
            Assert.Equal("_top", hyperlink.Anchor);
        }
    }
}


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
    }
}
