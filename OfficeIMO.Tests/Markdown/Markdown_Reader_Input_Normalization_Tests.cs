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
}
