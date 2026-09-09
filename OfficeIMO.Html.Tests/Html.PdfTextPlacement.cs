using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlPdfTextPlacement {
    [Theory]
    [InlineData("Arial")]
    [InlineData("Times New Roman")]
    [InlineData("Courier New")]
    public void PositionedCharactersUseTheWrittenFontsAdvance(string family) {
        byte[] bytes = HtmlConversionDocument.Parse(
            $"<p style='font-family:{family};font-size:24px;font-weight:bold;letter-spacing:1px'>NORTHWIND</p>")
            .ToPdfBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("NORTHWIND", string.Concat(letters.Select(letter => letter.Value)));
        for (int index = 1; index < letters.Length; index++) {
            double gap = letters[index].StartBaseLine.X - letters[index - 1].EndBaseLine.X;
            Assert.InRange(gap, 0.70D, 0.80D);
        }
    }

    [Theory]
    [InlineData("<strong>46.75 hours recorded.</strong> Three entries are approved.", "46.75 hours recorded. Three entries are approved.")]
    [InlineData("Confirm reference <strong>00130</strong>.", "Confirm reference 00130.")]
    [InlineData("<span>Leading</span> <em>and</em> <strong>trailing</strong> spaces.", "Leading and trailing spaces.")]
    public void InlineBoundariesPreserveSpacesAndDoNotOverlap(string body, string expected) {
        byte[] bytes = HtmlConversionDocument.Parse($"<p style='font:14px Arial'>{body}</p>").ToPdfBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.ToArray();
        Assert.Equal(expected, string.Concat(letters.Select(letter => letter.Value)));
        for (int index = 1; index < letters.Length; index++) {
            Assert.True(letters[index].StartBaseLine.X >= letters[index - 1].EndBaseLine.X - 0.02D,
                $"'{letters[index - 1].Value}' overlaps '{letters[index].Value}'.");
        }
    }

    [Fact]
    public void PositionedTextRetainsStandardFontPunctuation() {
        byte[] bytes = HtmlConversionDocument.Parse("<p>September · Approved • € £</p>").ToPdfBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains("September · Approved • € £", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }
}
