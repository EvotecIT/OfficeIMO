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
    public void Autolinks_Trim_Trailing_Bang_And_Question() {
        var doc = MarkdownReader.Parse("See https://example.com! And https://example.com?");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://example.com\">https://example.com</a>!", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"https://example.com\">https://example.com</a>?", html, StringComparison.Ordinal);
        Assert.DoesNotContain("https://example.com!</a>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("https://example.com?</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Keep_Balanced_Parentheses_In_Http_Urls() {
        var doc = MarkdownReader.Parse("See https://en.wikipedia.org/wiki/Function_(mathematics) and continue.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains(
            "<a href=\"https://en.wikipedia.org/wiki/Function_(mathematics)\">https://en.wikipedia.org/wiki/Function_(mathematics)</a>",
            html,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Trim_Unmatched_Closing_Paren_After_Http_Url() {
        var doc = MarkdownReader.Parse("See (https://en.wikipedia.org/wiki/Function_(mathematics)) now.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains(
            "(<a href=\"https://en.wikipedia.org/wiki/Function_(mathematics)\">https://en.wikipedia.org/wiki/Function_(mathematics)</a>)",
            html,
            StringComparison.Ordinal);
        Assert.DoesNotContain("mathematics))</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Ambiguous_Paren_Suffixed_Urls() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path_(x)).");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/path_(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/path_(x)).</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Www_Inside_Text() {
        var doc = MarkdownReader.Parse("See www.example.com, thanks.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://www.example.com\">www.example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Keep_Balanced_Parentheses_In_Www_Urls() {
        var doc = MarkdownReader.Parse("See www.example.com/path_(demo) next.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains(
            "<a href=\"https://www.example.com/path_(demo)\">www.example.com/path_(demo)</a>",
            html,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Email_Inside_Text() {
        var doc = MarkdownReader.Parse("Email user@example.com.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"mailto:user@example.com\">user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Mailto_Are_Supported() {
        var doc = MarkdownReader.Parse("Contact <mailto:user@example.com>.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"mailto:user@example.com\">mailto:user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Mailto_Respect_Url_Policy() {
        var options = new MarkdownReaderOptions {
            RestrictUrlSchemes = true,
            AllowedUrlSchemes = new[] { "http", "https" }
        };
        var doc = MarkdownReader.Parse("Contact <mailto:user@example.com>.", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("&lt;mailto:user@example.com&gt;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Absolute_Uris_Are_Supported() {
        var doc = MarkdownReader.Parse("Fetch <ftp://example.com/file.txt>.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"ftp://example.com/file.txt\">ftp://example.com/file.txt</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Tel_Uris_Are_Supported() {
        var doc = MarkdownReader.Parse("Call <tel:+123456789>.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"tel:+123456789\">tel:+123456789</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Urn_Uris_Are_Supported() {
        var doc = MarkdownReader.Parse("Lookup <urn:isbn:9780143127741>.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"urn:isbn:9780143127741\">urn:isbn:9780143127741</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_Absolute_Uris_Respect_Url_Policy() {
        var options = new MarkdownReaderOptions {
            RestrictUrlSchemes = true,
            AllowedUrlSchemes = new[] { "http", "https" }
        };
        var doc = MarkdownReader.Parse("Fetch <ftp://example.com/file.txt>.", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"ftp://example.com/file.txt\"", html, StringComparison.Ordinal);
        Assert.Contains("&lt;ftp://example.com/file.txt&gt;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Angle_Autolinks_Explicit_NonHierarchical_Uris_Respect_Url_Policy() {
        var options = new MarkdownReaderOptions {
            RestrictUrlSchemes = true,
            AllowedUrlSchemes = new[] { "http", "https" }
        };
        var doc = MarkdownReader.Parse("Call <tel:+123456789>.", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"tel:+123456789\"", html, StringComparison.Ordinal);
        Assert.Contains("&lt;tel:+123456789&gt;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Work_In_Tables_And_Lists() {
        var md = """
| Link |
| --- |
| www.example.com |

- Email user@example.com
""";
        var doc = MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"https://www.example.com\">www.example.com</a>", html, StringComparison.Ordinal);
        Assert.Contains("<a href=\"mailto:user@example.com\">user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Can_Be_Disabled() {
        var options = new MarkdownReaderOptions {
            AutolinkUrls = false,
            AutolinkWwwUrls = false,
            AutolinkEmails = false
        };
        var doc = MarkdownReader.Parse("See https://example.com and www.example.com and user@example.com", options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Require_Left_Boundary() {
        var doc = MarkdownReader.Parse("prefixhttps://example.com should not linkify.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
    }
}
