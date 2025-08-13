        [Fact]
        public void Html_Hyperlinks_IdAnchor() {
            string html = "<p id=\"section1\">Start</p><p><a href=\"#section1\" title=\"Back\" target=\"_blank\">Back</a></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Contains(doc.Bookmarks, b => b.Name == "section1");

            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;

            Assert.NotNull(hyperlink);
            Assert.Equal("Back", hyperlink.Tooltip);
            Assert.Equal(TargetFrame._blank, hyperlink.TargetFrame);
            Assert.Equal("section1", hyperlink.Anchor);
        }

        [Fact]
        public void Html_Hyperlinks_NameAnchor() {
            string html = "<p><a name=\"intro\"></a>Intro</p><p><a href=\"#intro\">Back</a></p>";

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
