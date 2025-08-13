using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

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
    public void Html_Hyperlinks_InternalAnchor_Enabled() {
        string html = "<p id=\"top\">Top</p><p><a href=\"#top\" title=\"Back\" target=\"_blank\">Back</a></p>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions { SupportsAnchorLinks = true });

        Assert.Contains(doc.Bookmarks, b => b.Name == "top");

        var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;

        Assert.NotNull(hyperlink);
        Assert.Equal("Back", hyperlink.Tooltip);
        Assert.Equal(TargetFrame._blank, hyperlink.TargetFrame);
        Assert.Equal("top", hyperlink.Anchor);
    }

    [Fact]
    public void Html_Hyperlinks_InternalAnchor_Disabled() {
        string html = "<p id=\"top\">Top</p><p><a href=\"#top\" title=\"Back\" target=\"_blank\">Back</a></p>";

        var doc = html.LoadFromHtml(new HtmlToWordOptions { SupportsAnchorLinks = false });

        Assert.Contains(doc.Bookmarks, b => b.Name == "top");
        Assert.Empty(doc.ParagraphsHyperLinks);
    }
}
            var hyperlink = doc.ParagraphsHyperLinks[0].Hyperlink;

            Assert.NotNull(hyperlink);
            Assert.Equal("Back", hyperlink.Tooltip);
            Assert.Equal(TargetFrame._blank, hyperlink.TargetFrame);
            Assert.Equal("top", hyperlink.Anchor);
        }
    }
}
