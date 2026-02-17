using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Input_Normalization_Tests {
    [Fact]
    public void Reader_Can_Normalize_SoftWrapped_Strong_BeforeParsing() {
        var options = new MarkdownReaderOptions {
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeSoftWrappedStrongSpans = true
            }
        };

        var html = MarkdownReader.Parse("**Status\nHEALTHY**", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<strong>Status HEALTHY</strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reader_Can_Normalize_InlineCode_LineBreaks_BeforeParsing() {
        var options = new MarkdownReaderOptions {
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeInlineCodeSpanLineBreaks = true
            }
        };

        var html = MarkdownReader.Parse("`a\nb`", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<code>a b</code>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reader_Can_Normalize_EscapedInlineCode_Via_Ast() {
        var options = new MarkdownReaderOptions {
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeEscapedInlineCodeSpans = true
            }
        };

        var html = MarkdownReader.Parse(@"Use \`/act act_001\` now.", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<code>/act act_001</code>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reader_Can_Normalize_TightStrongBoundaries_Via_Ast() {
        var options = new MarkdownReaderOptions {
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeTightStrongBoundaries = true
            }
        };

        var html = MarkdownReader.Parse("Status **Healthy**next", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<strong>Healthy</strong> next", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reader_Can_Normalize_LooseStrongDelimiters_BeforeParsing() {
        var options = new MarkdownReaderOptions {
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeLooseStrongDelimiters = true
            }
        };

        var html = MarkdownReader.Parse("check ** LDAP/Kerberos health on all DCs** next", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<strong>LDAP/Kerberos health on all DCs</strong> next", html, StringComparison.Ordinal);
        Assert.DoesNotContain("** LDAP/Kerberos health on all DCs**", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reader_Ast_Normalization_Propagates_To_Nested_Quote_Parsing() {
        var options = new MarkdownReaderOptions {
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeEscapedInlineCodeSpans = true,
                NormalizeTightStrongBoundaries = true
            }
        };

        var html = MarkdownReader.Parse("> Use \\`/act act_001\\` and **Healthy**next", options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<code>/act act_001</code>", html, StringComparison.Ordinal);
        Assert.Contains("<strong>Healthy</strong> next", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reader_Ast_Normalization_DoesNot_Change_Fenced_Code_Block_Content() {
        var options = new MarkdownReaderOptions {
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeEscapedInlineCodeSpans = true,
                NormalizeTightStrongBoundaries = true
            }
        };

        var markdown = """
```text
Use \`/act act_001\`
Status **Healthy**next
```
""";

        var parsed = MarkdownReader.Parse(markdown, options).ToMarkdown().Replace("\r\n", "\n");

        Assert.Contains("Use \\`/act act_001\\`", parsed, StringComparison.Ordinal);
        Assert.Contains("Status **Healthy**next", parsed, StringComparison.Ordinal);
    }

    [Fact]
    public void Reader_Can_Normalize_OrderedListMarkerSpacing_BeforeParsing() {
        var options = new MarkdownReaderOptions {
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeOrderedListMarkerSpacing = true,
                NormalizeLooseStrongDelimiters = true
            }
        };

        var markdown = """
1. **Privilege hygiene sweep**(Domain Admins + other privileged groups)
2.** Delegation risk audit**(unconstrained / constrained / protocol transition)
3.** Replication + DC health snapshot** (stale links, failing partners)
""";

        var html = MarkdownReader.Parse(markdown, options)
            .ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Equal(3, Count(html, "<li"));
        Assert.Contains("<strong>Delegation risk audit</strong>", html, StringComparison.Ordinal);
        Assert.Contains("<strong>Replication + DC health snapshot</strong>", html, StringComparison.Ordinal);
    }

    private static int Count(string value, string token) {
        if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(token)) return 0;

        int index = 0;
        int count = 0;
        while ((index = value.IndexOf(token, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += token.Length;
        }

        return count;
    }
}
