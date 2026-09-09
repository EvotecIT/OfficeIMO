using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlTableTextInheritance {
    [Theory]
    [InlineData(TableCaptionPosition.Above)]
    [InlineData(TableCaptionPosition.Below)]
    public void WordTableTextInheritsContainerGroupRowAndCellStyles(TableCaptionPosition captionPosition) {
        const string html = """
            <body style="background-color: #f2f5f7"><div style="color: #123456; font-family: Arial">
              <table style="font-size: 18px">
                <caption style="color: #147d78; font-weight: bold">Usage</caption>
                <thead style="color: #ffffff"><tr style="font-weight: bold">
                  <th style="background-color: #16324f">Hours</th>
                </tr></thead>
                <tbody><tr style="font-style: italic"><td>
                  Plain <span style="color: #a12030; font-style: normal">override</span>
                  <p>Block</p>
                  <table><tr><td>Nested</td></tr></table>
                </td></tr></tbody>
              </table>
            </div></body>
            """;
        using var document = HtmlConversionDocument.Parse(html).ToWordDocument(
            new HtmlToWordOptions { TableCaptionPosition = captionPosition });
        using var artifact = document.ToStream();
        using var package = WordprocessingDocument.Open(artifact, false);
        var errors = new OpenXmlValidator().Validate(package).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description + " " + error.Path?.XPath + " " + error.Node?.Parent?.OuterXml)));
        var body = package.MainDocumentPart!.Document.Body!;
        Run FindRun(string text) => Assert.Single(body.Descendants<Run>(), run => run.InnerText.Trim() == text);

        var header = FindRun("Hours");
        Assert.Equal("FFFFFF", header.RunProperties?.Color?.Val?.Value);
        Assert.NotNull(header.RunProperties?.Bold);
        Assert.Null(header.RunProperties?.Shading);
        Assert.Equal("16324F", header.Ancestors<TableCell>().First().TableCellProperties?.Shading?.Fill?.Value);
        Assert.Equal("147D78", FindRun("Usage").RunProperties?.Color?.Val?.Value);
        Assert.NotNull(FindRun("Usage").RunProperties?.Bold);
        Assert.Equal("A12030", FindRun("override").RunProperties?.Color?.Val?.Value);
        Assert.True(FindRun("override").RunProperties?.Italic == null || FindRun("override").RunProperties!.Italic!.Val?.Value == false);
        foreach (string text in new[] { "Plain", "Block", "Nested" }) {
            var run = FindRun(text);
            Assert.Equal("123456", run.RunProperties?.Color?.Val?.Value);
            Assert.NotNull(run.RunProperties?.Italic);
            Assert.Null(run.RunProperties?.Shading);
            Assert.Equal("Arial", run.RunProperties?.RunFonts?.Ascii?.Value);
        }
    }
}
