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
}
