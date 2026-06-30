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

    public static IEnumerable<object[]> CjkFriendlyEmphasisExtensionCases() {
        yield return new object[] { "japanese-closing-punctuation-strong", "これは**強調？**です" };
        yield return new object[] { "chinese-code-span-strong", "我可以强调**这个`code`**吗（Can I emphasize **this `code`**）？" };
        yield return new object[] { "cjk-neighbor-with-latin-parentheses-strong", "漢**(abc)**字" };
        yield return new object[] { "cjk-neighbor-with-fullwidth-punctuation-emphasis", "漢*（abc）*字" };
        yield return new object[] { "cjk-underscore-remains-literal", "漢__（abc）__字" };
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

    public static IEnumerable<object[]> GenericAttributesAbbreviationExtensionCases() {
        yield return new object[] { "abbreviation-inline", "*[HTML]: Hyper Text Markup Language\n\nHTML{#abbr .wide}" };
    }

    public static IEnumerable<object[]> GenericAttributesExtensionCases() {
        yield return new object[] { "atx-heading-id-class-title", "# Heading {#intro .wide title=\"Overview\"}" };
        yield return new object[] { "atx-heading-hash-suffix-attribute", "# C# {#intro .wide}" };
        yield return new object[] { "atx-heading-closing-marker-before-attribute", "# Heading # {#intro .wide}" };
        yield return new object[] { "atx-heading-attribute-before-closing-marker", "# Heading {#intro .wide} #" };
        yield return new object[] { "setext-heading-id-class-title", "Heading {#intro .wide title=\"Overview\"}\n=======" };
        yield return new object[] { "paragraph-id-class-title", "Paragraph {#intro .wide title=\"Overview\"}" };
        yield return new object[] { "paragraph-continuation-standalone-attribute-is-consumed", "Paragraph\n{#intro .wide}" };
        yield return new object[] { "paragraph-hardbreak-continuation-standalone-attribute-is-consumed", "Paragraph  \n{#intro .wide}" };
        yield return new object[] { "paragraph-backslash-hardbreak-continuation-standalone-attribute-is-consumed", "Paragraph\\\n{#intro .wide}" };
        yield return new object[] { "plain-text-no-space-attribute", "word{#plain .wide}" };
        yield return new object[] { "plain-text-plus-ending-no-space-attribute", "C++{#plain .wide}" };
        yield return new object[] { "plain-text-trailing-backtick-no-space-attribute", "text`{#plain .wide}" };
        yield return new object[] { "plain-text-trailing-double-backtick-no-space-attribute", "text``{#plain .wide}" };
        yield return new object[] { "plain-text-only-backtick-no-space-attribute", "`{#plain .wide}" };
        yield return new object[] { "plain-text-only-double-backtick-no-space-attribute", "``{#plain .wide}" };
        yield return new object[] { "escaped-star-no-space-attribute", "\\*{#esc .wide}" };
        yield return new object[] { "escaped-underscore-no-space-attribute", "\\_{#esc .wide}" };
        yield return new object[] { "escaped-backtick-no-space-attribute", "\\`{#esc .wide}" };
        yield return new object[] { "escaped-closing-paren-no-space-attribute", "\\){#esc .wide}" };
        yield return new object[] { "escaped-closing-bracket-no-space-attribute", "\\]{#esc .wide}" };
        yield return new object[] { "named-entity-no-space-attribute-stays-literal", "&copy;{#e .wide}" };
        yield return new object[] { "decimal-entity-no-space-attribute-stays-literal", "&#42;{#e .wide}" };
        yield return new object[] { "hex-entity-no-space-attribute-stays-literal", "&#x2A;{#e .wide}" };
        yield return new object[] { "thematic-break-like-line-is-attributed-paragraph", "--- {#hr .wide}" };
        yield return new object[] { "asterisk-thematic-break-like-line-is-attributed-paragraph", "*** {#hr .wide}" };
        yield return new object[] { "underscore-thematic-break-like-line-is-attributed-paragraph", "___ {#hr .wide}" };
        yield return new object[] { "standalone-attribute-before-heading", "{#intro .wide}\n# Heading" };
        yield return new object[] { "standalone-attribute-before-paragraph", "{#intro .wide}\nParagraph" };
        yield return new object[] { "standalone-attribute-before-fenced-code", "{#code .wide}\n```cs\nvar x = 1;\n```" };
        yield return new object[] { "list-contained-standalone-attribute-before-fenced-code", "- item\n\n  {#code .wide}\n  ```cs\n  x\n  ```" };
        yield return new object[] { "blockquote-contained-standalone-attribute-before-fenced-code", "> {#code .wide}\n> ```cs\n> x\n> ```" };
        yield return new object[] { "fenced-code-attribute-only-info-string", "```{#code .wide}\nvar x = 1;\n```" };
        yield return new object[] { "fenced-code-language-info-attributes", "```cs {#code .wide}\nvar x = 1;\n```" };
        yield return new object[] { "fenced-code-opaque-info-before-attributes", "```cs linenums {#code .wide}\nvar x = 1;\n```" };
        yield return new object[] { "tilde-fenced-code-attribute-only-info-string", "~~~{#code .wide}\nvar x = 1;\n~~~" };
        yield return new object[] { "standalone-attribute-before-unordered-list", "{#list .wide}\n- item" };
        yield return new object[] { "standalone-attribute-before-ordered-list", "{#list .wide}\n1. item" };
        yield return new object[] { "unordered-list-heading-attribute-stays-literal", "- # Heading {#h .wide}" };
        yield return new object[] { "ordered-list-heading-attribute-stays-literal", "1. # Heading {#h .wide}" };
        yield return new object[] { "loose-list-nested-heading-attribute-stays-literal", "- item\n\n  # Heading {#h .wide}" };
        yield return new object[] { "standalone-attribute-before-inline-image-paragraph", "{#img .wide}\n![Alt](image.png)" };
        yield return new object[] { "standalone-attribute-before-blockquote-stays-literal", "{#q .wide}\n> quote" };
        yield return new object[] { "standalone-attribute-before-html-block", "{#html .wide}\n<div>raw</div>" };
        yield return new object[] { "standalone-attribute-before-reference-definition", "{#ref .wide}\n[id]: https://example.com\n\n[site][id]" };
        yield return new object[] { "standalone-attribute-before-thematic-break", "{#rule .wide}\n---" };
        yield return new object[] { "standalone-attribute-before-indented-code", "{#code .wide}\n    var x = 1;" };
        yield return new object[] { "standalone-attribute-before-definition-looking-text", "{#term .wide}\nTerm\n: definition" };
        yield return new object[] { "blockquote-paragraph-attribute-stays-literal", "> quote {#q .lead}" };
        yield return new object[] { "blockquote-atx-heading-attribute-stays-literal", "> # Heading {#h .wide}" };
        yield return new object[] { "blockquote-setext-heading-attribute-stays-literal", "> Heading {#h .wide}\n> -------" };
        yield return new object[] { "nested-blockquote-paragraph-attribute-stays-literal", "> > quote {#q .lead}" };
        yield return new object[] { "blockquote-standalone-attribute-before-unordered-list", "> {#list .wide}\n> - item" };
        yield return new object[] { "blockquote-standalone-attribute-before-ordered-list", "> {#list .wide}\n> 1. item" };
        yield return new object[] { "unordered-list-item-attribute-is-consumed", "- item {#li .selected}" };
        yield return new object[] { "ordered-list-item-attribute-is-consumed", "1. item {#li .selected}" };
        yield return new object[] { "blockquote-list-item-attribute-is-consumed", "> - item {#li .selected}" };
        yield return new object[] { "nested-list-item-attribute-is-consumed", "- outer\n  - inner {#li .selected}" };
        yield return new object[] { "inline-link-id-class-title", "[site](https://example.com){#lnk .primary title=\"Site\"}" };
        yield return new object[] { "inline-link-label-attribute-promotes-to-paragraph", "[site{#txt .wide}](https://example.com)" };
        yield return new object[] { "inline-emphasis-id-class", "*emphasis*{#em .marked}" };
        yield return new object[] { "inline-emphasis-content-attribute-promotes-to-paragraph", "*em{#inner .wide}*" };
        yield return new object[] { "inline-strong-id-class", "**strong**{#strong .marked}" };
        yield return new object[] { "inline-strong-content-attribute-promotes-to-paragraph", "**strong{#inner .wide}**" };
        yield return new object[] { "inline-strong-emphasis-id-class", "***both***{#both .mix}" };
        yield return new object[] { "inline-code-id-class", "`code`{#code .token}" };
        yield return new object[] { "inline-image-id-class", "![alt](img.png){#img .wide}" };
        yield return new object[] { "inline-image-alt-attribute-promotes-to-paragraph", "![alt{#alt .wide}](img.png)" };
        yield return new object[] { "linked-image-id-class", "[![alt](img.png)](https://example.com){#linked .wide}" };
        yield return new object[] { "linked-image-alt-attribute-promotes-to-paragraph", "[![alt{#alt .wide}](img.png)](https://example.com)" };
        yield return new object[] { "inline-html-span-attribute-stays-literal", "<span>hi</span>{#span .wide}" };
    }

    public static IEnumerable<object[]> GenericAttributesReferenceExtensionCases() {
        yield return new object[] { "full-reference-link", "[site][id]{#lnk .primary}\n\n[id]: https://example.com" };
        yield return new object[] { "collapsed-reference-link", "[site][]{#lnk .primary}\n\n[site]: https://example.com" };
        yield return new object[] { "shortcut-reference-link", "[site]{#lnk .primary}\n\n[site]: https://example.com" };
        yield return new object[] { "full-reference-image", "![alt][img]{#img .wide}\n\n[img]: img.png" };
        yield return new object[] { "collapsed-reference-image", "![img][]{#img .wide}\n\n[img]: img.png" };
        yield return new object[] { "shortcut-reference-image", "![img]{#img .wide}\n\n[img]: img.png" };
    }

    public static IEnumerable<object[]> GenericAttributesAutoLinksExtensionCases() {
        yield return new object[] { "bare-url-paragraph-attribute", "https://example.com{#auto .wide}" };
        yield return new object[] { "angle-url-autolink", "<https://example.com>{#auto .wide}" };
        yield return new object[] { "angle-email-autolink", "<user@example.com>{#mail .wide}" };
    }

    public static IEnumerable<object[]> GenericAttributesEmphasisExtrasExtensionCases() {
        yield return new object[] { "strikethrough", "~~gone~~{#s .strike}" };
        yield return new object[] { "highlight", "==mark=={#m .mark}" };
        yield return new object[] { "inserted", "++ins++{#i .insert}" };
        yield return new object[] { "superscript", "^sup^{#sup .high}" };
        yield return new object[] { "subscript", "~sub~{#sub .low}" };
    }

    public static IEnumerable<object[]> GenericAttributesPipeTableExtensionCases() {
        yield return new object[] { "standalone-attribute-before-pipe-table", "{#tbl .wide title=\"Overview\"}\n| A |\n|---|\n| B |" };
        yield return new object[] { "pipe-table-header-cell-attribute", "| A {#tbl .wide title=\"Overview\"} |\n|---|\n| B |" };
        yield return new object[] { "pipe-table-second-header-cell-attribute", "| A | B {#tbl .wide} |\n|---|---|\n| C | D |" };
        yield return new object[] { "pipe-table-body-cell-attribute", "| A |\n|---|\n| B {#tbl .wide} |" };
        yield return new object[] { "blockquote-pipe-table-cell-attribute", "> | A {#tbl .wide} |\n> |---|\n> | B |" };
    }

    public static IEnumerable<object[]> GenericAttributesTaskListExtensionCases() {
        yield return new object[] { "blockquote-standalone-attribute-before-task-list", "> {#list .wide}\n> - [x] task" };
        yield return new object[] { "task-list-item-attribute-is-consumed", "- [x] task {#task .done}" };
        yield return new object[] { "unchecked-task-list-item-attribute-is-consumed", "- [ ] task {#task .todo}" };
        yield return new object[] { "nested-task-list-item-attribute-is-consumed", "- outer\n  - [x] task {#task .done}" };
    }

    public static IEnumerable<object[]> GenericAttributesFootnoteExtensionCases() {
        yield return new object[] { "footnote-definition-paragraph-attribute", "[^a]: note {#fn .wide}\n\ntext[^a]" };
        yield return new object[] { "footnote-second-paragraph-attribute-stays-literal", "[^a]: first\n\n    second {#p .wide}\n\ntext[^a]" };
        yield return new object[] { "standalone-attribute-before-footnote-definition", "{#fn .wide}\n[^a]: note\n\ntext[^a]" };
    }

    public static IEnumerable<object[]> GenericAttributesFootnoteReferenceExtensionCases() {
        yield return new object[] { "footnote-reference-attribute-is-consumed", "text[^a]{#ref .wide}\n\n[^a]: note" };
    }

    public static IEnumerable<object[]> GenericAttributesDefinitionListExtensionCases() {
        yield return new object[] { "definition-list-definition-paragraph-attribute", "Term\n:   Definition {#def .wide}" };
        yield return new object[] { "definition-list-second-paragraph-attribute-stays-literal", "Term\n:   first\n\n    second {#p .wide}" };
        yield return new object[] { "definition-list-term-attribute", "Term {#term .wide}\n:   Definition" };
    }

    public static IEnumerable<object[]> AlertBlocksExtensionCases() {
        yield return new object[] { "note-paragraph", "> [!NOTE]\n> Body" };
        yield return new object[] { "note-empty", "> [!NOTE]" };
        yield return new object[] { "lowercase-note", "> [!note]\n> Body" };
        yield return new object[] { "note-followed-by-outside-paragraph", "> [!NOTE]\nOutside" };
        yield return new object[] { "tip-rich-inline-body", "> [!TIP]\n> Use **strong** [links](https://example.com)." };
        yield return new object[] { "important-multiple-paragraphs", "> [!IMPORTANT]\n> First paragraph\n>\n> Second paragraph" };
        yield return new object[] { "warning-list", "> [!WARNING]\n> - Item" };
        yield return new object[] { "caution-fenced-code", "> [!CAUTION]\n> ```ps1\n> Get-Item .\n> ```" };
        yield return new object[] { "note-nested-blockquote", "> [!NOTE]\n> > Nested quote" };
        yield return new object[] { "custom-kind", "> [!CUSTOM]\n> Body" };
        yield return new object[] { "numeric-marker-stays-blockquote", "> [!NOTE1]\n> Body" };
        yield return new object[] { "hyphen-marker-stays-blockquote", "> [!NOTE-TIP]\n> Body" };
        yield return new object[] { "note-title-stays-blockquote", "> [!NOTE] Title\n> Body" };
        yield return new object[] { "note-strong-title-stays-blockquote", "> [!NOTE] **Title**\n> Body" };
        yield return new object[] { "custom-title-stays-blockquote", "> [!CUSTOM] Title\n> Body" };
        yield return new object[] { "github-docs-note", "> [!NOTE]\n> Useful information that users should know, even when skimming content." };
        yield return new object[] { "github-docs-tip", "> [!TIP]\n> Helpful advice for doing things better or more easily." };
        yield return new object[] { "github-docs-important", "> [!IMPORTANT]\n> Key information users need to know to achieve their goal." };
        yield return new object[] { "github-docs-warning", "> [!WARNING]\n> Urgent info that needs immediate user attention to avoid problems." };
        yield return new object[] { "github-docs-caution", "> [!CAUTION]\n> Advises about risks or negative outcomes of certain actions." };
        yield return new object[] { "github-docs-five-alerts-separated", "> [!NOTE]\n> Useful information that users should know, even when skimming content.\n\n> [!TIP]\n> Helpful advice for doing things better or more easily.\n\n> [!IMPORTANT]\n> Key information users need to know to achieve their goal.\n\n> [!WARNING]\n> Urgent info that needs immediate user attention to avoid problems.\n\n> [!CAUTION]\n> Advises about risks or negative outcomes of certain actions." };
        yield return new object[] { "alert-between-paragraphs", "Before\n\n> [!NOTE]\n> Body\n\nAfter" };
        yield return new object[] { "alert-inside-list-item-boundary", "- item\n  > [!NOTE]\n  > Body" };
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
    [MemberData(nameof(AlertBlocksExtensionCases))]
    public void MarkdownReader_AlertBlocks_Match_Markdig_Extension(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        MarkdownBlockRenderBuiltInExtensions.AddMarkdigAlertHtmlFallback(htmlOptions);
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseAlertBlocks(builder, null);

        var officeOptions = new MarkdownReaderOptions {
            CalloutTitleMode = MarkdownCalloutTitleMode.MarkdigCompatible
        };

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeAlertHtmlForParity(markdig), NormalizeAlertHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(AlertBlocksExtensionCases))]
    public void MarkdownReader_AlertBlocks_Writer_Reparse_Matches_Markdig_Extension(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        MarkdownBlockRenderBuiltInExtensions.AddMarkdigAlertHtmlFallback(htmlOptions);
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseAlertBlocks(builder, null);

        var officeOptions = new MarkdownReaderOptions {
            CalloutTitleMode = MarkdownCalloutTitleMode.MarkdigCompatible
        };

        var written = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" });
        var officeReparsed = MarkdownReader
            .Parse(written, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeAlertHtmlForParity(markdig), NormalizeAlertHtmlForParity(officeReparsed));
    }

    [Theory]
    [MemberData(nameof(AutoLinksExtensionCases))]
    public void MarkdownReader_GfmAutolinks_Match_Markdig_AutoLinks_Extension(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            PercentEncodeTildeInUrlAttributes = true,
            EscapeNonAsciiText = false
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
            PercentEncodeTildeInUrlAttributes = true,
            EscapeNonAsciiText = false
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
    [MemberData(nameof(CjkFriendlyEmphasisExtensionCases))]
    public void MarkdownReader_CjkFriendlyEmphasis_Match_Markdig_Extension(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseCjkFriendlyEmphasis(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.CjkFriendlyEmphasis = true;

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

    [Theory]
    [MemberData(nameof(GenericAttributesAbbreviationExtensionCases))]
    public void MarkdownReader_GenericAttributes_On_Abbreviations_Match_Markdig_Extensions(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseAbbreviations(builder);
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.Abbreviations = true;
        officeOptions.GenericAttributes = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeGenericAttributesHtmlForParity(markdig), NormalizeGenericAttributesHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(GenericAttributesExtensionCases))]
    public void MarkdownReader_GenericAttributes_On_Blocks_Match_Markdig_Extension(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.GenericAttributes = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeGenericAttributesHtmlForParity(markdig), NormalizeGenericAttributesHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(GenericAttributesReferenceExtensionCases))]
    public void MarkdownReader_GenericAttributes_On_Reference_Links_And_Images_Match_Markdig_Extension(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.GenericAttributes = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeGenericAttributesHtmlForParity(markdig), NormalizeGenericAttributesHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(GenericAttributesAutoLinksExtensionCases))]
    public void MarkdownReader_GenericAttributes_On_Autolinks_Match_Markdig_Extensions(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false,
            PercentEncodeTildeInUrlAttributes = true
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseAutoLinks(builder);
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        officeOptions.GenericAttributes = true;
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

        Assert.Equal(NormalizeGenericAttributesHtmlForParity(markdig), NormalizeGenericAttributesHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(GenericAttributesEmphasisExtrasExtensionCases))]
    public void MarkdownReader_GenericAttributes_On_EmphasisExtras_Match_Markdig_Extensions(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseEmphasisExtras(builder);
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.GenericAttributes = true;
        officeOptions.Subscript = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeGenericAttributesHtmlForParity(markdig), NormalizeGenericAttributesHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(GenericAttributesPipeTableExtensionCases))]
    public void MarkdownReader_GenericAttributes_In_PipeTables_Match_Markdig_Extensions(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UsePipeTables(builder);
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.Tables = true;
        officeOptions.GenericAttributes = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeGenericAttributesHtmlForParity(markdig), NormalizeGenericAttributesHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(GenericAttributesTaskListExtensionCases))]
    public void MarkdownReader_GenericAttributes_In_TaskLists_Match_Markdig_Extensions(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseTaskLists(builder);
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        officeOptions.GenericAttributes = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeTaskListHtmlForParity(markdig), NormalizeTaskListHtmlForParity(office));
    }

    [Theory]
    [MemberData(nameof(GenericAttributesFootnoteExtensionCases))]
    public void MarkdownReader_GenericAttributes_In_Footnotes_Match_Markdig_Extensions(string caseName, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false,
            GitHubFootnoteHtml = true
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseFootnotes(builder);
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        officeOptions.GenericAttributes = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        var normalizedOffice = NormalizeGenericAttributesHtmlForParity(office);
        var normalizedMarkdig = NormalizeGenericAttributesHtmlForParity(markdig);

        if (string.Equals(caseName, "standalone-attribute-before-footnote-definition", StringComparison.Ordinal)) {
            Assert.DoesNotContain("{#fn .wide}", normalizedMarkdig, StringComparison.Ordinal);
            Assert.DoesNotContain("{#fn .wide}", normalizedOffice, StringComparison.Ordinal);
            Assert.DoesNotContain("id=\"fn\" class=\"wide\"", normalizedMarkdig, StringComparison.Ordinal);
            Assert.DoesNotContain("id=\"fn\" class=\"wide\"", normalizedOffice, StringComparison.Ordinal);
            Assert.DoesNotContain("<p id=\"fn\" class=\"wide\"></p>", normalizedOffice, StringComparison.Ordinal);
            return;
        }

        if (string.Equals(caseName, "footnote-second-paragraph-attribute-stays-literal", StringComparison.Ordinal)) {
            Assert.Contains("second {#p .wide}", normalizedMarkdig, StringComparison.Ordinal);
            Assert.DoesNotContain("id=\"p\" class=\"wide\"", normalizedMarkdig, StringComparison.Ordinal);
            Assert.Contains("second {#p .wide}", normalizedOffice, StringComparison.Ordinal);
            Assert.DoesNotContain("id=\"p\" class=\"wide\"", normalizedOffice, StringComparison.Ordinal);
            return;
        }

        Assert.Contains("<p id=\"fn\" class=\"wide\">note ", normalizedMarkdig, StringComparison.Ordinal);
        Assert.Contains("<p id=\"fn\" class=\"wide\">note ", normalizedOffice, StringComparison.Ordinal);
        Assert.DoesNotContain("{#fn .wide}", normalizedOffice, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(GenericAttributesFootnoteReferenceExtensionCases))]
    public void MarkdownReader_GenericAttributes_After_FootnoteReferences_Match_Markdig_Extensions(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false,
            GitHubFootnoteHtml = true
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseFootnotes(builder);
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        officeOptions.GenericAttributes = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        var normalizedOffice = NormalizeGenericAttributesHtmlForParity(office);
        var normalizedMarkdig = NormalizeGenericAttributesHtmlForParity(markdig);

        Assert.DoesNotContain("{#ref .wide}", normalizedMarkdig, StringComparison.Ordinal);
        Assert.DoesNotContain("{#ref .wide}", normalizedOffice, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(GenericAttributesDefinitionListExtensionCases))]
    public void MarkdownReader_GenericAttributes_In_DefinitionLists_Match_Markdig_Extensions(string _, string markdown) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        var builder = new Markdig.MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UseDefinitionLists(builder);
        Markdig.MarkdownExtensions.UseGenericAttributes(builder);

        var officeOptions = MarkdownReaderOptions.CreatePortableProfile();
        officeOptions.DefinitionLists = true;
        officeOptions.GenericAttributes = true;

        var office = MarkdownReader
            .Parse(markdown, officeOptions)
            .ToHtmlFragment(htmlOptions);
        var markdig = MarkdigMarkdown.ToHtml(markdown, builder.Build());

        Assert.Equal(NormalizeGenericAttributesHtmlForParity(markdig), NormalizeGenericAttributesHtmlForParity(office));
    }

    private static string NormalizeAlertHtmlForParity(string html) {
        return Regex.Replace(
            NormalizeHtmlForParity(html),
            "<svg[\\s\\S]*?</svg>",
            "<svg />",
            RegexOptions.CultureInvariant);
    }

    private static string NormalizeGenericAttributesHtmlForParity(string html) {
        var normalized = NormalizeHtmlForParity(html);
        normalized = Regex.Replace(
            normalized,
            "(<h[1-6][^>]*>[^<]*?)\\s+</h",
            "$1</h",
            RegexOptions.CultureInvariant);
        return NormalizeImageAttributeOrderForParity(normalized);
    }

    private static string NormalizeImageAttributeOrderForParity(string html) {
        return Regex.Replace(
            html,
            "<img\\s+([^>]*?)\\s*/>",
            match => {
                var attributes = Regex.Matches(match.Groups[1].Value, "([A-Za-z_:][-A-Za-z0-9_:.]*)(?:=\"([^\"]*)\")?");
                var ordered = attributes
                    .Cast<Match>()
                    .Select(static attr => new KeyValuePair<string, string?>(attr.Groups[1].Value, attr.Groups[2].Success ? attr.Groups[2].Value : null))
                    .OrderBy(static attr => GetImageAttributeOrder(attr.Key))
                    .ThenBy(static attr => attr.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(static attr => attr.Value == null ? attr.Key : attr.Key + "=\"" + attr.Value + "\"");
                return "<img " + string.Join(" ", ordered) + " />";
            },
            RegexOptions.CultureInvariant);
    }

    private static int GetImageAttributeOrder(string attributeName) {
        return attributeName.ToLowerInvariant() switch {
            "src" => 0,
            "id" => 1,
            "class" => 2,
            "alt" => 3,
            "title" => 4,
            _ => 100
        };
    }

    private static string NormalizeTaskListHtmlForParity(string html) {
        var normalized = NormalizeGenericAttributesHtmlForParity(html)
            .Replace("checked=\"\"", "checked=\"checked\"")
            .Replace("disabled=\"\"", "disabled=\"disabled\"");
        normalized = Regex.Replace(
            normalized,
            "<input[^>]*>",
            match => match.Value.Contains("checked", StringComparison.OrdinalIgnoreCase)
                ? "<input type=\"checkbox\" checked=\"checked\" />"
                : "<input type=\"checkbox\" />",
            RegexOptions.CultureInvariant);
        return normalized;
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
