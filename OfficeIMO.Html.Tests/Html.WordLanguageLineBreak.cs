using OfficeIMO.Html;
using OfficeIMO.Word.Html;
using System.Linq;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlWordLanguageLineBreakTests {
    [Theory]
    [InlineData("\n")]
    [InlineData("\n\n")]
    public void HtmlImportPreservesLanguageTaggedBreakOnlyText(string breaks) {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<html lang='en'><body><p>Text<span style='white-space:pre'>" + breaks + "</span>More</p></body></html>");

        using var document = source.ToWordDocumentResult().RequireValue();

        var insertedBreaks = document.Paragraphs.Where(paragraph => paragraph.IsBreak).ToArray();
        Assert.Equal(breaks.Length, insertedBreaks.Length);
        Assert.All(insertedBreaks, paragraph => Assert.Equal("en", paragraph.Language));
        string exportedHtml = document.ToHtml();
        Assert.Contains("Text", exportedHtml);
        Assert.Contains("More", exportedHtml);
        Assert.Equal(breaks.Length, Regex.Matches(exportedHtml, "<br\\b", RegexOptions.IgnoreCase).Count);
    }
}
