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
    public void Autolinks_DoNot_Link_Balanced_Paren_Urls_With_Trailing_Comma() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path_(x), ok");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/path_(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/path_(x), ok</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Balanced_Paren_Urls_With_Trailing_Dot() {
        var doc = MarkdownReader.Parse("Visit www.example.com/path_(x).");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com/path_(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit www.example.com/path_(x).</p>", html, StringComparison.Ordinal);
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
    public void Autolinks_DoNot_Link_Http_Urls_With_Query_Parentheses() {
        var doc = MarkdownReader.Parse("Visit https://example.com/search?q=(x) now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/search?q=(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/search?q=(x) now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Urls_With_Query_Parentheses() {
        var doc = MarkdownReader.Parse("Visit www.example.com/search?q=(x) now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com/search?q=(x)\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit www.example.com/search?q=(x) now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_Still_Link_Path_Parentheses_Before_Query_String() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path_(demo)?q=value now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains(
            "<a href=\"https://example.com/path_(demo)?q=value\">https://example.com/path_(demo)?q=value</a>",
            html,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Http_Urls_With_Query_Ampersands() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path?q=1&next=2 now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/path?q=1&amp;next=2\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/path?q=1&amp;next=2 now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Http_Urls_With_Fragment_Ampersands() {
        var doc = MarkdownReader.Parse("Visit https://example.com/path#frag&next now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com/path#frag&amp;next\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit https://example.com/path#frag&amp;next now</p>", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Visit _https://example.com now")]
    [InlineData("Visit /https://example.com now")]
    public void Autolinks_DoNot_Link_Http_Urls_After_Invalid_Left_Boundaries(string markdown) {
        var doc = MarkdownReader.Parse(markdown);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        Assert.Contains(markdown.Replace("&", "&amp;"), html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Visit foo:https://example.com now")]
    [InlineData("Visit foo.https://example.com now")]
    public void Autolinks_DoNot_Link_Http_Urls_After_Colon_Or_Dot(string markdown) {
        var doc = MarkdownReader.Parse(markdown);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://example.com\"", html, StringComparison.Ordinal);
        Assert.Contains(markdown, html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Urls_With_Query_Ampersands() {
        var doc = MarkdownReader.Parse("Visit www.example.com/path?q=1&next=2 now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com/path?q=1&amp;next=2\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit www.example.com/path?q=1&amp;next=2 now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Urls_After_Invalid_Left_Boundaries() {
        var doc = MarkdownReader.Parse("Visit _www.example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit _www.example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Www_Urls_After_Colon() {
        var doc = MarkdownReader.Parse("Visit foo:www.example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"https://www.example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Visit foo:www.example.com now</p>", html, StringComparison.Ordinal);
    }


    [Fact]
    public void Autolinks_Email_Inside_Text() {
        var doc = MarkdownReader.Parse("Email user@example.com.");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<a href=\"mailto:user@example.com\">user@example.com</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Mailto_Email_Tokens() {
        var doc = MarkdownReader.Parse("Contact mailto:user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact mailto:user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_After_Slash() {
        var doc = MarkdownReader.Parse("Contact /user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact /user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_After_Colon() {
        var doc = MarkdownReader.Parse("Contact foo:user@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact foo:user@example.com now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_With_Path_Suffixes() {
        var doc = MarkdownReader.Parse("Contact user@example.com/path now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact user@example.com/path now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_With_Fragment_Suffixes() {
        var doc = MarkdownReader.Parse("Contact user@example.com#frag now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact user@example.com#frag now</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Autolinks_DoNot_Link_Plain_Emails_With_Plus_Tags() {
        var doc = MarkdownReader.Parse("Contact user.name+tag@example.com now");
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("href=\"mailto:user.name+tag@example.com\"", html, StringComparison.Ordinal);
        Assert.Contains("<p>Contact user.name+tag@example.com now</p>", html, StringComparison.Ordinal);
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
