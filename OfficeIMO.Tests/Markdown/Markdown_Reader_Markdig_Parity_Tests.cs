using OfficeIMO.Markdown;
using MarkdigMarkdown = Markdig.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Markdig_Parity_Tests {
    public static IEnumerable<object[]> CoreParityCases() {
        yield return new object[] { "italic-with-inner-bold", "*a **b** c*" };
        yield return new object[] { "bold-with-inner-italic", "**a *b* c**" };
        yield return new object[] { "blockquote-lazy-continuation", "> Quote line 1\nQuote line 2" };
        yield return new object[] { "blockquote-blank-line", "> Quote\n>\n> Continued line" };
        yield return new object[] { "blockquote-nested-list", "> - List item\n>   - Nested" };
    }

    [Theory]
    [MemberData(nameof(CoreParityCases))]
    public void MarkdownReader_Matches_Markdig_On_Curated_Cases(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };

        var office = MarkdownReader.Parse(markdown).ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown);

        Assert.Equal(NormalizeHtmlForParity(markdig), NormalizeHtmlForParity(office));
    }

    private static string NormalizeHtmlForParity(string html) {
        if (string.IsNullOrWhiteSpace(html)) return string.Empty;

        var sb = new StringBuilder(html.Length);
        bool inTag = false;
        bool lastWasWhitespace = false;

        for (int i = 0; i < html.Length; i++) {
            char ch = html[i];
            if (ch == '<') {
                if (!inTag && lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                    sb.Append(' ');
                }

                inTag = true;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (ch == '>') {
                inTag = false;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (inTag) {
                sb.Append(ch);
                continue;
            }

            if (char.IsWhiteSpace(ch)) {
                lastWasWhitespace = true;
                continue;
            }

            if (lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                sb.Append(' ');
            }

            lastWasWhitespace = false;
            sb.Append(ch);
        }

        var normalized = sb.ToString().Replace("> <", "><", StringComparison.Ordinal);
        normalized = normalized
            .Replace(" <ul", "<ul", StringComparison.Ordinal)
            .Replace(" <ol", "<ol", StringComparison.Ordinal)
            .Replace(" <blockquote", "<blockquote", StringComparison.Ordinal)
            .Replace(" <pre", "<pre", StringComparison.Ordinal)
            .Replace(" <table", "<table", StringComparison.Ordinal)
            .Replace(" <p", "<p", StringComparison.Ordinal);

        return normalized.Trim();
    }
}
