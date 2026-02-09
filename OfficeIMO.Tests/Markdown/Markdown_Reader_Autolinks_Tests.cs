using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Autolinks_Tests {
    [Fact]
    public void Autolinks_Http_Inside_Text() {
        var doc = MarkdownReader.Parse("See https://example.com.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://example.com\">https://example.com</a>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("https://example.com.</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Www_Inside_Text() {
        var doc = MarkdownReader.Parse("See www.example.com, thanks.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://www.example.com\">www.example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Email_Inside_Text() {
        var doc = MarkdownReader.Parse("Email user@example.com.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"mailto:user@example.com\">user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Require_Left_Boundary() {
        var doc = MarkdownReader.Parse("prefixhttps://example.com should not linkify.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
    }
}

