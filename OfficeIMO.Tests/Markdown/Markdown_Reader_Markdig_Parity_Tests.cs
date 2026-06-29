using OfficeIMO.Markdown;
using MarkdigMarkdown = Markdig.Markdown;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Markdig_Parity_Tests {
    public static IEnumerable<object[]> CoreParityCases() {
        yield return new object[] { "italic-with-inner-bold", "*a **b** c*" };
        yield return new object[] { "bold-with-inner-italic", "**a *b* c**" };
        yield return new object[] { "triple-marker-balanced", "***foo***" };
        yield return new object[] { "triple-marker-inner-bold-then-outer-italic", "***foo** bar*" };
        yield return new object[] { "single-star-inside-bold-stays-literal", "**foo*bar**" };
        yield return new object[] { "double-star-inside-italic-rebalances", "*a **b* c**" };
        yield return new object[] { "double-star-opener-degrades-to-literal-and-italic", "**foo*" };
        yield return new object[] { "quad-star-opener-degrades-to-literal-and-triple", "****foo***" };
        yield return new object[] { "quad-underscore-opener-degrades-to-literal-and-triple", "____foo___" };
        yield return new object[] { "sextuple-star-opener-degrades-to-literal-and-triple", "******foo***" };
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
        yield return new object[] { "unordered-list-becomes-loose-when-later-item-has-second-paragraph", "- a\n- b\n\n  second paragraph" };
        yield return new object[] { "ordered-list-becomes-loose-when-later-item-has-second-paragraph", "10. a\n11. b\n\n    second paragraph" };
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
        yield return new object[] { "autolink-query-balanced-parens", "Visit https://example.com/search?q=(x) now" };
        yield return new object[] { "autolink-query-balanced-parens-then-dot", "Visit https://example.com/search?q=(x)." };
        yield return new object[] { "autolink-query-balanced-parens-then-comma", "Visit https://example.com/search?q=(x), now" };
        yield return new object[] { "autolink-www-query-balanced-parens", "Visit www.example.com/search?q=(x) now" };
        yield return new object[] { "autolink-fragment-balanced-parens", "Visit https://example.com/path#(x) now" };
        yield return new object[] { "autolink-www-fragment-balanced-parens", "Visit www.example.com/path#(x) now" };
        yield return new object[] { "autolink-fragment-with-ampersand", "Visit https://example.com/path#frag&next now" };
        yield return new object[] { "autolink-www-query-with-ampersand", "Visit www.example.com/path?q=1&next=2 now" };
        yield return new object[] { "autolink-after-underscore", "Visit _https://example.com now" };
        yield return new object[] { "autolink-www-after-underscore", "Visit _www.example.com now" };
        yield return new object[] { "autolink-after-slash", "Visit /https://example.com now" };
        yield return new object[] { "autolink-after-colon", "Visit foo:https://example.com now" };
        yield return new object[] { "autolink-www-after-colon", "Visit foo:www.example.com now" };
        yield return new object[] { "autolink-after-dot", "Visit foo.https://example.com now" };
        yield return new object[] { "autolink-after-plus", "Visit foo+https://example.com now" };
        yield return new object[] { "autolink-www-after-plus", "Visit foo+www.example.com now" };
        yield return new object[] { "autolink-after-dash", "Visit foo-https://example.com now" };
        yield return new object[] { "autolink-www-after-dash", "Visit foo-www.example.com now" };
        yield return new object[] { "autolink-after-equals", "Visit foo=https://example.com now" };
        yield return new object[] { "autolink-www-after-equals", "Visit foo=www.example.com now" };
        yield return new object[] { "autolink-after-ampersand", "Visit &https://example.com now" };
        yield return new object[] { "autolink-www-after-ampersand", "Visit &www.example.com now" };
        yield return new object[] { "autolink-after-open-paren", "Visit (https://example.com now" };
        yield return new object[] { "autolink-after-open-paren-with-close", "Visit (https://example.com) now" };
        yield return new object[] { "autolink-www-after-open-paren", "Visit (www.example.com now" };
        yield return new object[] { "autolink-www-after-open-paren-with-close", "Visit (www.example.com) now" };
        yield return new object[] { "autolink-after-apostrophe", "Visit 'https://example.com now" };
        yield return new object[] { "autolink-www-after-apostrophe", "Visit 'www.example.com now" };
        yield return new object[] { "autolink-after-open-bracket", "Visit [https://example.com now" };
        yield return new object[] { "autolink-www-after-open-bracket", "Visit [www.example.com now" };
        yield return new object[] { "angle-autolink-http", "<https://example.com>" };
        yield return new object[] { "plain-mailto-does-not-autolink-email", "Contact mailto:user@example.com now" };
        yield return new object[] { "plain-email-after-colon", "Contact foo:user@example.com now" };
        yield return new object[] { "plain-email-after-equals", "Contact foo=user@example.com now" };
        yield return new object[] { "plain-email-after-ampersand", "Contact &user@example.com now" };
        yield return new object[] { "plain-email-after-open-paren", "Contact (user@example.com) now" };
        yield return new object[] { "plain-email-after-apostrophe", "Contact 'user@example.com now" };
        yield return new object[] { "plain-email-after-open-bracket", "Contact [user@example.com now" };
        yield return new object[] { "plain-email-after-slash", "Contact /user@example.com now" };
        yield return new object[] { "plain-email-after-underscore", "Contact _user@example.com now" };
        yield return new object[] { "plain-email-with-path-suffix", "Contact user@example.com/path now" };
        yield return new object[] { "plain-email-with-fragment-suffix", "Contact user@example.com#frag now" };
        yield return new object[] { "quote-blank-paragraph-then-paragraph", "> one\n>\n> \n> two" };
        yield return new object[] { "unordered-list-lazy-continuation", "- item\ncontinuation" };
        yield return new object[] { "ordered-list-lazy-continuation", "1. item\ncontinuation" };
        yield return new object[] { "unordered-list-loose-lazy-continuation-after-indented-paragraph", "- item\n\n    code\nafter" };
        yield return new object[] { "unordered-list-indented-code-then-paragraph", "- item\n\n      code\n\n  after" };
        yield return new object[] { "ordered-list-nested-blockquote-then-code", "1. item\n   > quote\n\n      code" };
        yield return new object[] { "setext-heading-before-list", "Heading\n-------\n- item" };
        yield return new object[] { "blockquote-setext-heading", "> title\n> -----" };
        yield return new object[] { "blockquote-setext-heading-then-paragraph", "> title\n> -----\n>\n> after" };
        yield return new object[] { "list-setext-heading", "- item\n  heading\n  -------" };
        yield return new object[] { "list-setext-heading-then-paragraph-same-group", "- item\n  heading\n  -------\n  after" };
        yield return new object[] { "list-blank-line-then-setext-heading", "- item\n\n  Heading\n  ---\n  text" };
        yield return new object[] { "list-setext-heading-then-quote", "- item\n  heading\n  -------\n\n  > quote" };
        yield return new object[] { "paragraph-then-nonone-ordered-marker", "alpha\n10. beta" };
        yield return new object[] { "list-continuation-then-nonone-ordered-marker", "- outer\n  10. item\n      continuation" };
        yield return new object[] { "list-quote-lazy-nonone-ordered-continued", "- outer\n  > alpha\n  10. beta\n      gamma" };
        yield return new object[] { "blockquote-heading-then-list", "> Heading\n> -------\n>\n> 1. item" };
        yield return new object[] { "blockquote-heading-then-nonone-list-text", "> Heading\n> -------\n>\n> 10. item" };
        yield return new object[] { "nonone-ordered-marker-with-indented-continuation", "alpha\n10. beta\n    gamma" };
        yield return new object[] { "list-quote-lazy-after-setext-heading", "- outer\n  heading\n  -------\n  > quote\n  continuation" };
        yield return new object[] { "literal-url-colon-stays-paragraph", "Visit https://example.com/path_(x): now" };
        yield return new object[] { "atx-empty-heading", "#" };
        yield return new object[] { "atx-indented-heading", "   # heading" };
        yield return new object[] { "atx-trailing-closing-hashes", "### foo ###" };
        yield return new object[] { "inline-link-nested-label", "[link [inner]](https://example.com)" };
        yield return new object[] { "inline-link-escaped-closing-paren", "[x](https://example.com/a\\)b)" };
        yield return new object[] { "inline-image-balanced-parens", "Look ![alt](https://example.com/a_(b).png) now" };
        yield return new object[] { "linked-image-nested-alt", "[![alt [x]](https://example.com/a_(b).png)](https://example.com)" };
        yield return new object[] { "inline-link-angle-destination-space", "[x](<https://example.com/a b> \"title\")" };
        yield return new object[] { "inline-image-angle-destination-space", "Look ![x](<https://example.com/a b> \"title\") now" };
        yield return new object[] { "reference-link-angle-destination-space", "[x][r]\n\n[r]: <https://example.com/a b>" };
        yield return new object[] { "reference-link-first-definition-wins", "[x][r]\n\n[r]: https://first.example.com \"first\"\n[r]: https://second.example.com \"second\"" };
        yield return new object[] { "reference-link-first-definition-wins-with-next-line-title", "[x][r]\n\n[r]: https://first.example.com\n  \"first\"\n[r]: https://second.example.com \"second\"" };
        yield return new object[] { "inline-link-empty-angle-destination", "[x](<>)" };
        yield return new object[] { "inline-link-empty-angle-destination-with-title", "[x](<> \"title\")" };
        yield return new object[] { "inline-image-empty-angle-destination", "Look ![x](<>) now" };
        yield return new object[] { "html-block-type6-continues-until-blank-line", "<div>\ninner\n</div>\nParagraph" };
        yield return new object[] { "html-block-type7-continues-until-blank-line", "<widget-box>\ninner\n</widget-box>\nParagraph" };
        yield return new object[] { "inline-link-invalid-title-tail", "[x](https://example.com \"title\" extra)" };
        yield return new object[] { "inline-link-title-with-escaped-quote", "[x](https://example.com \"a \\\"quote\\\" title\")" };
        yield return new object[] { "reference-link-empty-angle-destination", "[x][r]\n\n[r]: <>" };
        yield return new object[] { "reference-link-title-next-line", "[x][r]\n\n[r]: https://example.com\n  \"title\"" };
        yield return new object[] { "reference-link-angle-destination-title-next-line", "[x][r]\n\n[r]: <https://example.com/a b>\n  \"title\"" };
        yield return new object[] { "reference-link-label-whitespace-normalization", "[x][A   B]\n\n[a b]: https://example.com" };
        yield return new object[] { "reference-link-label-case-normalization", "[x][MiXeD]\n\n[mixed]: https://example.com" };
        yield return new object[] { "reference-link-escaped-bracket-label", "[x][a \\[b\\]]\n\n[a \\[b\\]]: https://example.com" };
        yield return new object[] { "reference-link-collapsed", "[x][]\n\n[x]: https://example.com" };
        yield return new object[] { "reference-link-invalid-nested-label-definition", "[x [y]]\n\n[x [y]]: https://example.com" };
        yield return new object[] { "reference-link-invalid-empty-label-definition", "[]: https://example.com" };
        yield return new object[] { "reference-link-invalid-whitespace-label-definition", "[ ]: https://example.com" };
        yield return new object[] { "reference-link-invalid-title-tail", "[x]: https://example.com \"title\" extra" };
        yield return new object[] { "reference-link-invalid-angle-title-tail", "[x]: <https://example.com/a b> \"title\" extra" };
        yield return new object[] { "reference-link-invalid-extra-body-token", "[x]: https://example.com extra" };
        yield return new object[] { "reference-link-invalid-missing-angle-close", "[x]: <https://example.com/a b" };
        yield return new object[] { "reference-link-invalid-unclosed-title", "[x]: https://example.com \"title" };
        yield return new object[] { "reference-link-invalid-next-line-title-tail", "[x]: https://example.com\n  \"title\" extra" };
        yield return new object[] { "reference-link-invalid-angle-next-line-title-tail", "[x]: <https://example.com/a b>\n  \"title\" extra" };
        yield return new object[] { "reference-link-shortcut", "[x]\n\n[x]: https://example.com" };
        yield return new object[] { "reference-link-definition-three-space-indent", "[x][r]\n\n   [r]: https://example.com" };
        yield return new object[] { "reference-link-definition-tab-indent-invalid", "[x][r]\n\n\t[r]: https://example.com" };
        yield return new object[] { "unordered-list-tab-continuation", "- first line\n\tsecond line\n- next" };
        yield return new object[] { "ordered-list-tab-continuation", "1. first line\n\tsecond line\n2. next" };
        yield return new object[] { "fenced-code-open-indent-four-is-indented-code", "    ```csharp\n    var x = 1;\n    ```" };
        yield return new object[] { "fenced-code-close-indent-four-does-not-close", "```csharp\nvar x = 1;\n    ```\nafter" };
        yield return new object[] { "backtick-fence-info-string-cannot-contain-backtick", "``` c`sharp\nbody\n```" };
        yield return new object[] { "fenced-code-brace-metadata-keeps-primary-language-html", "```chart {#summary .wide title=\"Quarterly Revenue\"}\nbody\n```" };
        yield return new object[] { "fenced-code-malformed-brace-metadata-keeps-primary-language-html", "```chart {#summary .wide title=\"Quarterly Revenue\"\nbody\n```" };
        yield return new object[] { "blockquote-lazy-after-unordered-list-item", "> - item\ncontinuation" };
        yield return new object[] { "blockquote-lazy-after-ordered-list-item", "> 1. item\ncontinuation" };
        yield return new object[] { "blockquote-explicit-after-ordered-list-item", "> 1. item\n>   continuation" };
        yield return new object[] { "blockquote-indented-paragraph-then-lazy-continuation", "> quote\n>     code\ncontinuation" };
        yield return new object[] { "nested-blockquote-lazy-after-list-item", "> Outer\n> > Inner\n> > - a\n> > - b\n> After" };
    }

    public static IEnumerable<object[]> PortableProfileCases() {
        yield return new object[] { "bare-http-default-markdig", "Visit https://example.com now" };
        yield return new object[] { "bare-www-default-markdig", "Visit www.example.com now" };
        yield return new object[] { "bare-email-default-markdig", "Contact user@example.com now" };
        yield return new object[] { "mixed-literal-autolinks-default-markdig", "See https://example.com and www.example.com and user@example.com" };
        yield return new object[] { "angle-email-still-autolinks", "Email <user@example.com>." };
        yield return new object[] { "angle-mailto-still-autolinks", "Contact <mailto:user@example.com> now" };
        yield return new object[] { "literal-http-after-colon", "Visit foo:https://example.com now" };
        yield return new object[] { "literal-http-after-open-paren", "Visit (https://example.com) now" };
        yield return new object[] { "literal-http-after-apostrophe", "Visit 'https://example.com now" };
        yield return new object[] { "literal-http-after-open-bracket", "Visit [https://example.com now" };
        yield return new object[] { "literal-email-after-colon", "Contact foo:user@example.com now" };
        yield return new object[] { "literal-email-after-open-paren", "Contact (user@example.com) now" };
        yield return new object[] { "literal-email-after-apostrophe", "Contact 'user@example.com now" };
        yield return new object[] { "literal-email-after-open-bracket", "Contact [user@example.com now" };
        yield return new object[] { "literal-email-with-plus-tag", "Contact user.name+tag@example.com now" };
        yield return new object[] { "callout-stays-blockquote-text", "> [!NOTE]\n> body\ntext" };
        yield return new object[] { "callout-with-blank-line-stays-blockquote-text", "> [!NOTE]\n>\n> body\ntext" };
        yield return new object[] { "unordered-task-stays-plain-list-text", "- [ ] task\n  continuation" };
        yield return new object[] { "ordered-task-stays-plain-list-text", "1. [x] task\ncontinuation" };
    }

    public static IEnumerable<object[]> AutoLinksExtensionCases() {
        yield return new object[] { "http-query-ampersand", "Visit https://example.com/path?q=1&next=2 now" };
        yield return new object[] { "http-fragment-ampersand", "Visit https://example.com/path#frag&next now" };
        yield return new object[] { "www-query-ampersand", "Visit www.example.com/path?q=1&next=2 now" };
        yield return new object[] { "http-query-parens", "Visit https://example.com/search?q=(x) now" };
        yield return new object[] { "www-query-parens", "Visit www.example.com/search?q=(x) now" };
        yield return new object[] { "http-balanced-parens-extra-close", "Visit https://example.com/path_(x)). now" };
        yield return new object[] { "www-balanced-parens-extra-close", "Visit www.example.com/path_(x)). now" };
        yield return new object[] { "https-trailing-period-before-close-paren", "Visit https://example.com/path.) now" };
        yield return new object[] { "https-trailing-comma-before-close-paren", "Visit https://example.com/path,) now" };
        yield return new object[] { "https-trailing-semicolon-before-close-paren", "Visit https://example.com/path;) now" };
        yield return new object[] { "https-trailing-bang-before-close-paren", "Visit https://example.com/path!) now" };
        yield return new object[] { "https-trailing-question-before-close-paren", "Visit https://example.com/path?) now" };
        yield return new object[] { "www-trailing-period-before-close-paren", "Visit www.example.com/path.) now" };
        yield return new object[] { "https-trailing-semicolon-links", "Visit https://example.com/path; now" };
        yield return new object[] { "https-trailing-double-semicolon-links", "Visit https://example.com/path;; now" };
        yield return new object[] { "uppercase-www-prefix-stays-literal", "Visit WWW.example.com now" };
        yield return new object[] { "mixed-case-www-host", "Visit www.Example.com now" };
        yield return new object[] { "unicode-http-domain", "Visit https://пример.рф/path now" };
        yield return new object[] { "unicode-http-path", "Visit https://example.com/ścieżka?q=zażółć now" };
        yield return new object[] { "http-url-path-tilde", "Visit https://example.com/path~tilde now" };
        yield return new object[] { "http-url-closing-bracket", "Visit https://example.com/path] now" };
        yield return new object[] { "http-url-trailing-single-quote", "Visit https://example.com/path' now" };
        yield return new object[] { "http-url-trailing-double-quote", "Visit https://example.com/path\" now" };
        yield return new object[] { "http-url-paired-single-quotes-stays-literal", "Visit 'https://example.com/path' now" };
        yield return new object[] { "http-url-userinfo-stays-literal", "Visit https://user@example.com/path now" };
        yield return new object[] { "http-url-host-underscore-stays-literal", "Visit https://exa_mple.com/path now" };
        yield return new object[] { "www-url-path-tilde", "Visit www.example.com/path~tilde now" };
        yield return new object[] { "www-url-trailing-semicolon-links", "Visit www.example.com/path; now" };
        yield return new object[] { "www-url-userinfo-stays-literal", "Visit www.user@example.com/path now" };
        yield return new object[] { "unicode-www-domain", "Visit www.пример.рф/path now" };
        yield return new object[] { "unicode-ftp-domain", "Visit ftp://пример.рф/path now" };
        yield return new object[] { "ftp-url", "Visit ftp://example.com/file.txt now" };
        yield return new object[] { "ftp-url-trailing-semicolon-links", "Visit ftp://example.com/file; now" };
        yield return new object[] { "ftp-url-userinfo-stays-literal", "Visit ftp://user@example.com/file now" };
        yield return new object[] { "ftp-url-host-underscore-stays-literal", "Visit ftp://exa_mple.com/file now" };
        yield return new object[] { "ftp-url-query-ampersand", "Visit ftp://example.com/path?q=1&next=2 now" };
        yield return new object[] { "ftp-url-query-parens", "Visit ftp://example.com/search?q=(x) now" };
        yield return new object[] { "ftp-url-trailing-dot", "Visit ftp://example.com/file.txt. now" };
        yield return new object[] { "http-url-trailing-double-dot", "Visit https://example.com/path.. now" };
        yield return new object[] { "http-url-trailing-underscore", "Visit https://example.com/path_ now" };
        yield return new object[] { "http-url-trailing-double-underscore", "Visit https://example.com/path__ now" };
        yield return new object[] { "www-url-trailing-underscore", "Visit www.example.com_ now" };
        yield return new object[] { "www-url-host-underscore-stays-literal", "Visit www.exa_mple.com now" };
        yield return new object[] { "http-url-after-underscore", "Visit _https://example.com now" };
        yield return new object[] { "ftp-localhost-stays-literal", "Visit ftp://localhost/file now" };
        yield return new object[] { "tel-url", "Call tel:+123456789 now" };
        yield return new object[] { "tel-url-after-apostrophe-stays-literal", "Call 'tel:+123-456 now" };
        yield return new object[] { "tel-url-trailing-semicolon-links", "Call tel:+123-456; now" };
        yield return new object[] { "tel-url-trailing-dot", "Call tel:+123-456. now" };
        yield return new object[] { "tel-url-parentheses", "Call tel:(123)456 now" };
        yield return new object[] { "xmpp-url", "Chat xmpp:user@example.com now" };
        yield return new object[] { "uppercase-ftp-stays-literal", "Visit FTP://example.com/file now" };
        yield return new object[] { "uppercase-tel-stays-literal", "Call TEL:+123-456 now" };
        yield return new object[] { "lowercase-mailto-links", "Contact mailto:user@example.com now" };
        yield return new object[] { "lowercase-mailto-after-apostrophe-stays-literal", "Contact 'mailto:user@example.com now" };
        yield return new object[] { "lowercase-mailto-address-trailing-colon-links", "Contact mailto:user@example.com:: now" };
        yield return new object[] { "lowercase-mailto-address-trailing-dash-links", "Contact mailto:user@example.com- now" };
        yield return new object[] { "lowercase-mailto-address-semicolon-stays-literal", "Contact mailto:user@example.com; now" };
        yield return new object[] { "lowercase-mailto-path-links", "Contact mailto:user@example.com/path now" };
        yield return new object[] { "lowercase-mailto-path-semicolon-links", "Contact mailto:user@example.com/path; now" };
        yield return new object[] { "lowercase-mailto-path-query-links", "Contact mailto:user@example.com/path?q=1 now" };
        yield return new object[] { "lowercase-mailto-query-semicolon-links", "Contact mailto:user@example.com?subject=Hi; now" };
        yield return new object[] { "lowercase-mailto-path-trailing-underscore", "Contact mailto:user@example.com/path__ now" };
        yield return new object[] { "uppercase-mailto-stays-literal", "Contact MAILTO:user@example.com now" };
        yield return new object[] { "plain-email-stays-literal", "Contact user@example.com now" };
    }

    public static IEnumerable<object[]> AutoLinksPipeTableExtensionCases() {
        yield return new object[] { "table-http-query-ampersand", "| Link |\n| --- |\n| https://example.com/path?q=1&next=2 |\n" };
        yield return new object[] { "table-www-query-parens", "| Link |\n| --- |\n| www.example.com/search?q=(x) |\n" };
        yield return new object[] { "table-http-path-tilde", "| Link |\n| --- |\n| https://example.com/path~tilde |\n" };
        yield return new object[] { "table-mailto-url", "| Link |\n| --- |\n| mailto:user@example.com |\n" };
        yield return new object[] { "table-plain-email-stays-literal", "| Link |\n| --- |\n| user@example.com |\n" };
    }

    public static IEnumerable<object[]> EmphasisExtrasExtensionCases() {
        yield return new object[] { "inserted", "++inserted++" };
        yield return new object[] { "inserted-with-nested-emphasis", "++inserted *and emphasized*++" };
        yield return new object[] { "single-plus-stays-literal", "+inserted+" };
        yield return new object[] { "superscript", "2^10^" };
        yield return new object[] { "superscript-after-space", "x ^10^" };
        yield return new object[] { "superscript-at-start", "^alone^" };
        yield return new object[] { "nested-superscript", "2^^10^^" };
        yield return new object[] { "superscript-with-nested-emphasis", "^super *em*^" };
        yield return new object[] { "superscript-with-trailing-text", "2^10^tail" };
        yield return new object[] { "superscript-with-whitespace-before-close-stays-literal", "2^10 ^" };
        yield return new object[] { "subscript", "H~2~O" };
        yield return new object[] { "subscript-after-space", "x ~sub~" };
        yield return new object[] { "subscript-at-start", "~alone~" };
        yield return new object[] { "subscript-and-double-tilde-strike", "H~2~ and ~~del~~" };
        yield return new object[] { "subscript-with-nested-emphasis", "~sub *em*~" };
        yield return new object[] { "subscript-with-trailing-text", "H~2~tail" };
        yield return new object[] { "subscript-with-whitespace-before-close-stays-literal", "H~2 ~" };
    }

    public static IEnumerable<object[]> AbbreviationExtensionCases() {
        yield return new object[] { "basic", "*[HTML]: Hyper Text Markup Language\nHTML test" };
        yield return new object[] { "multiple-occurrences", "*[HTML]: Hyper Text Markup Language\nHTML and HTML." };
        yield return new object[] { "does-not-match-inside-word", "*[CSS]: Cascading Style Sheets\nCSS3 CSS CSS-like" };
        yield return new object[] { "multiple-definitions", "*[HTML]: Hyper Text Markup Language\n*[CSS]: Cascading Style Sheets\nHTML CSS" };
        yield return new object[] { "definition-after-earlier-paragraph", "HTML before\n\n*[HTML]: Hyper Text Markup Language" };
        yield return new object[] { "heading-inline", "*[HTML]: Hyper Text Markup Language\n\n# HTML heading" };
        yield return new object[] { "code-span-stays-literal", "*[HTML]: Hyper Text Markup Language\n\n`HTML` HTML" };
        yield return new object[] { "duplicate-last-definition-wins", "*[HTML]: First\n*[HTML]: Second\nHTML" };
        yield return new object[] { "case-sensitive", "*[html]: Lower\nHTML html Html" };
        yield return new object[] { "unicode-label", "*[åbc]: Unicode\nåbc ÅBC" };
        yield return new object[] { "punctuation-label", "*[C++]: Language\nC++ C+++ C++-like" };
        yield return new object[] { "trailing-dash-boundary", "*[HTML]: Hyper Text Markup Language\nHTML- HTML-like" };
        yield return new object[] { "opening-punctuation-boundaries-stay-literal", "*[HTML]: Hyper Text Markup Language\n(HTML) 'HTML' \"HTML\" /HTML .HTML" };
        yield return new object[] { "unresolved-bracket-text", "*[HTML]: Hyper Text Markup Language\n[HTML]" };
        yield return new object[] { "list-item-definition", "- *[HTML]: Hyper Text Markup Language\n- HTML" };
        yield return new object[] { "empty-title", "*[HTML]:   \nHTML" };
        yield return new object[] { "emphasis-inline", "*[HTML]: Hyper Text Markup Language\n\n*HTML* and **HTML**" };
        yield return new object[] { "link-label-inline", "*[HTML]: Hyper Text Markup Language\n\n[HTML](https://example.com)" };
        yield return new object[] { "blockquote-inline", "*[HTML]: Hyper Text Markup Language\n\n> HTML quoted" };
        yield return new object[] { "list-inline", "*[HTML]: Hyper Text Markup Language\n\n- HTML item" };
        yield return new object[] { "definition-later-applies-earlier", "HTML before\n\n*[HTML]: Hyper Text Markup Language\n\nHTML after" };
    }

    public static IEnumerable<object[]> AbbreviationPipeTableExtensionCases() {
        yield return new object[] { "table-cell-inline", "*[HTML]: Hyper Text Markup Language\n\n| Term |\n| --- |\n| HTML |" };
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

    [Theory]
    [MemberData(nameof(PortableProfileCases))]
    public void MarkdownReader_Matches_Markdig_With_Portable_Profile(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };

        var office = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreatePortableProfile()).ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown);

        Assert.Equal(NormalizeHtmlForParity(markdig), NormalizeHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(AutoLinksExtensionCases))]
    public void MarkdownReader_GfmAutolinks_Match_Markdig_AutoLinks_Extension(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            PercentEncodeTildeInUrlAttributes = true
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseAutoLinks(builder);

        var officeOptions = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        officeOptions.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis = true;
        officeOptions.AutolinkTrimSingleTrailingPunctuationOrUnderscore = true;
        officeOptions.AutolinkKeepTrailingSemicolonPunctuation = true;
        officeOptions.AutolinkRejectUnderscoreInWwwHost = true;
        officeOptions.AutolinkRejectUnderscoreInUrlHost = true;
        officeOptions.AutolinkRejectUserInfoAuthority = true;
        officeOptions.AutolinkAllowClosingBracketInUrl = true;
        officeOptions.AutolinkKeepTrailingQuotePunctuation = true;
        officeOptions.AutolinkEmails = false;
        officeOptions.AutolinkBareMailtoDisplayAddressOnly = true;
        officeOptions.AutolinkBareMailtoMarkdigSemicolonHandling = true;
        officeOptions.AutolinkValidPreviousCharacters = "_('";
        officeOptions.AutolinkBareSchemePrefixes = new[] { "mailto:", "ftp://", "tel:" };

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeHtmlForParity(markdig), NormalizeHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(AutoLinksPipeTableExtensionCases))]
    public void MarkdownReader_GfmAutolinks_In_PipeTables_Match_Markdig_Extensions(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            PercentEncodeTildeInUrlAttributes = true
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UsePipeTables(builder);
        Markdig.MarkdownExtensions.UseAutoLinks(builder);

        var officeOptions = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        officeOptions.AutolinkAllowTrailingPunctuationBeforeClosingParenthesis = true;
        officeOptions.AutolinkTrimSingleTrailingPunctuationOrUnderscore = true;
        officeOptions.AutolinkKeepTrailingSemicolonPunctuation = true;
        officeOptions.AutolinkRejectUnderscoreInWwwHost = true;
        officeOptions.AutolinkRejectUnderscoreInUrlHost = true;
        officeOptions.AutolinkRejectUserInfoAuthority = true;
        officeOptions.AutolinkAllowClosingBracketInUrl = true;
        officeOptions.AutolinkKeepTrailingQuotePunctuation = true;
        officeOptions.AutolinkEmails = false;
        officeOptions.AutolinkBareMailtoDisplayAddressOnly = true;
        officeOptions.AutolinkBareMailtoMarkdigSemicolonHandling = true;
        officeOptions.AutolinkValidPreviousCharacters = "_('";
        officeOptions.AutolinkBareSchemePrefixes = new[] { "mailto:", "ftp://", "tel:" };

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeHtmlForParity(markdig), NormalizeHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(EmphasisExtrasExtensionCases))]
    public void MarkdownReader_EmphasisExtras_Match_Markdig_EmphasisExtras_Extension(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseEmphasisExtras(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.Subscript = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeHtmlForParity(markdig), NormalizeHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(AbbreviationExtensionCases))]
    public void MarkdownReader_Abbreviations_Match_Markdig_Abbreviations_Extension(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseAbbreviations(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.Abbreviations = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeHtmlForParity(markdig), NormalizeHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(AbbreviationPipeTableExtensionCases))]
    public void MarkdownReader_Abbreviations_In_PipeTables_Match_Markdig_Extensions(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UsePipeTables(builder);
        Markdig.MarkdownExtensions.UseAbbreviations(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.Abbreviations = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

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

        var normalized = sb.ToString()
            .Replace("> <", "><")
            .Replace("&#39;", "'")
            .Replace("&#x27;", "'");
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
