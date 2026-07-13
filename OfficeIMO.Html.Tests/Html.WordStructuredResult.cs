using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlWordStructuredResult {
    [Fact]
    public void WordHtml_ResultCarriesOnlyCurrentConversionDiagnostics() {
        var options = new HtmlToWordOptions {
            UnsupportedCssHandling = HtmlUnsupportedCssHandling.Warn
        };

        HtmlToWordResult first = OfficeIMO.Html.HtmlConversionDocument.Parse("<p style='future-property:value'>First</p>").ToWordDocumentResult(options);
        HtmlToWordResult second = OfficeIMO.Html.HtmlConversionDocument.Parse("<p>Second</p>").ToWordDocumentResult(options);
        try {
            Assert.Contains(first.Report.Diagnostics, diagnostic => diagnostic.Code == "UnsupportedCssDeclaration");
            Assert.Empty(second.Report.Diagnostics);
            Assert.True(first.Succeeded);
            Assert.True(second.Succeeded);
        } finally {
            first.Value.Dispose();
            second.Value.Dispose();
        }
    }

    [Fact]
    public void WordHtml_UsesConversionScopedMissingStyleResolver() {
        var options = new HtmlToWordOptions {
            StyleMissingHandler = args => args.Style = WordParagraphStyles.Heading1
        };

        HtmlToWordResult result = OfficeIMO.Html.HtmlConversionDocument.Parse("<p class='custom-heading'>Heading</p>").ToWordDocumentResult(options);
        using WordDocument document = result.Value;

        Assert.Equal(WordParagraphStyles.Heading1, Assert.Single(document.Paragraphs).Style);
    }
}
