using OfficeIMO.Markdown;
using MarkdigMarkdown = Markdig.Markdown;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Markdig_Parity_Tests {
    public static IEnumerable<object[]> CoreParityCases() {
        yield return new object[] { "italic-with-inner-bold", "*a **b** c*" };
        yield return new object[] { "bold-with-inner-italic", "**a *b* c**" };
        yield return new object[] { "blockquote-lazy-continuation", "> Quote line 1\nQuote line 2" };
        yield return new object[] { "blockquote-blank-line", "> Quote\n>\n> Continued line" };
        yield return new object[] { "blockquote-nested-list", "> - List item\n>   - Nested" };
        yield return new object[] { "blockquote-indented-lazy-text", "> Quote line 1\n    indented continuation" };
        yield return new object[] { "blockquote-indented-lazy-listlike", "> Quote line 1\n    - nested text" };
        yield return new object[] { "list-quote-then-nested-list", "- item\n  > quote\n  continuation\n  - nested" };
        yield return new object[] { "list-quote-trailing-paragraph", "- item\n\n  > quote\n\n  trailing" };
        yield return new object[] { "intraword-underscore", "foo_bar_baz" };
        yield return new object[] { "quote-fenced-code", "> ```\n> code\n> ```" };
        yield return new object[] { "ordered-same-type-nesting", "1. item\n   1. nested" };
        yield return new object[] { "list-quote-then-fence", "- item\n  > quote\n\n    code" };
        yield return new object[] { "loose-list-followed-by-paragraph", "- item\n\n  second paragraph" };
        yield return new object[] { "autolink-trailing-punctuation", "Visit https://example.com/path_(x))." };
        yield return new object[] { "quoted-indented-code", "> para\n>\n>     code" };
        yield return new object[] { "quoted-blank-line-then-list", "> intro\n>\n> - item" };
        yield return new object[] { "list-item-fenced-code", "- item\n\n  ```js\n  let x = 1;\n  ```" };
        yield return new object[] { "blockquote-ends-before-nonquoted-paragraph", "> quote\n\noutside" };
        yield return new object[] { "blockquote-ends-before-nonquoted-list", "> quote\n- outside" };
        yield return new object[] { "list-blank-line-then-quote-then-paragraph", "- item\n\n  > quote\n\n  after" };
        yield return new object[] { "ordered-list-fenced-code-followed-by-paragraph", "1. item\n\n   ```txt\n   code\n   ```\n\n   after" };
        yield return new object[] { "autolink-balanced-parens-then-comma", "Visit https://example.com/path_(x), ok" };
        yield return new object[] { "autolink-www-balanced-parens-then-dot", "Visit www.example.com/path_(x)." };
        yield return new object[] { "angle-autolink-http", "<https://example.com>" };
        yield return new object[] { "quote-blank-paragraph-then-paragraph", "> one\n>\n> \n> two" };
        yield return new object[] { "unordered-list-indented-code-then-paragraph", "- item\n\n      code\n\n  after" };
        yield return new object[] { "ordered-list-nested-blockquote-then-code", "1. item\n   > quote\n\n      code" };
        yield return new object[] { "setext-heading-before-list", "Heading\n-------\n- item" };
        yield return new object[] { "blockquote-setext-heading", "> title\n> -----" };
        yield return new object[] { "blockquote-setext-heading-then-paragraph", "> title\n> -----\n>\n> after" };
        yield return new object[] { "list-setext-heading", "- item\n  heading\n  -------" };
        yield return new object[] { "list-setext-heading-then-quote", "- item\n  heading\n  -------\n\n  > quote" };
        yield return new object[] { "paragraph-then-nonone-ordered-marker", "alpha\n10. beta" };
        yield return new object[] { "list-continuation-then-nonone-ordered-marker", "- outer\n  10. item\n      continuation" };
        yield return new object[] { "list-quote-lazy-nonone-ordered-continued", "- outer\n  > alpha\n  10. beta\n      gamma" };
        yield return new object[] { "blockquote-heading-then-list", "> Heading\n> -------\n>\n> 1. item" };
        yield return new object[] { "blockquote-heading-then-nonone-list-text", "> Heading\n> -------\n>\n> 10. item" };
        yield return new object[] { "nonone-ordered-marker-with-indented-continuation", "alpha\n10. beta\n    gamma" };
        yield return new object[] { "list-quote-lazy-after-setext-heading", "- outer\n  heading\n  -------\n  > quote\n  continuation" };
        yield return new object[] { "literal-url-colon-stays-paragraph", "Visit https://example.com/path_(x): now" };
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

        var normalized = sb.ToString().Replace("> <", "><");
        normalized = Regex.Replace(normalized, "<h([1-6])\\s+id=\"[^\"]*\">", "<h$1>", RegexOptions.CultureInvariant);
        normalized = normalized
            .Replace(" <ul", "<ul")
            .Replace(" <ol", "<ol")
            .Replace(" <blockquote", "<blockquote")
            .Replace(" <pre", "<pre")
            .Replace(" <table", "<table")
            .Replace(" <p", "<p");

        return normalized.Trim();
    }
}
