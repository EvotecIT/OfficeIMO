using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlWordRunFontSizeTests {
    [Theory]
    [InlineData(10, "font-size:10pt")]
    [InlineData(10.5, "font-size:10.5pt")]
    public void WordToHtmlWritesRunFontSizeWithOneUnit(double points, string expected) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("SIZED").SetFontSize(OfficeLength.Points(points));

        string html = document.ToHtml();

        Assert.Contains(expected + "\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("ptpt", html, StringComparison.Ordinal);
    }
}
