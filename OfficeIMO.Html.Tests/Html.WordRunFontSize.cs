using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeIMO.Tests;

public sealed class HtmlWordRunFontSizeTests {
    [Theory]
    [InlineData(10, "font-size:10pt", false)]
    [InlineData(10.5, "font-size:10.5pt", false)]
    [InlineData(10, "font-size:10pt", true)]
    [InlineData(10.5, "font-size:10.5pt", true)]
    public void WordToHtmlWritesRunFontSizeWithOneUnit(double points, string expected, bool inEquationContainer) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("SIZED").SetFontSize(OfficeLength.Points(points));
        if (inEquationContainer) {
            Run run = paragraph._run!;
            run.Remove();
            paragraph._paragraph.Append(new SdtRun(
                new SdtProperties(),
                new SdtContentRun(run, new M.OfficeMath(new M.Run(new M.Text("x"))))));
        }

        string html = document.ToHtml();

        var parsed = new HtmlParser().ParseDocument(html);
        var sizedRun = Assert.Single(parsed.QuerySelectorAll("span"),
            span => span.TextContent == "SIZED" &&
                span.GetAttribute("style")?.StartsWith("font-size:", StringComparison.Ordinal) == true);
        Assert.Equal(expected, sizedRun.GetAttribute("style"));
        Assert.DoesNotContain("ptpt", html, StringComparison.Ordinal);
        if (inEquationContainer) Assert.NotNull(parsed.QuerySelector("math"));
    }
}
