using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordHighlightPaletteTests {
    // ECMA-376 Part 1 (5th edition), 17.18.40 ST_HighlightColor.
    [Theory]
    [InlineData("black", "000000")]
    [InlineData("blue", "0000FF")]
    [InlineData("cyan", "00FFFF")]
    [InlineData("darkBlue", "00008B")]
    [InlineData("darkCyan", "008B8B")]
    [InlineData("darkGray", "A9A9A9")]
    [InlineData("darkGreen", "006400")]
    [InlineData("darkMagenta", "800080")]
    [InlineData("darkRed", "8B0000")]
    [InlineData("darkYellow", "808000")]
    [InlineData("green", "00FF00")]
    [InlineData("lightGray", "D3D3D3")]
    [InlineData("magenta", "FF00FF")]
    [InlineData("red", "FF0000")]
    [InlineData("white", "FFFFFF")]
    [InlineData("yellow", "FFFF00")]
    public void InheritedHighlightHasTheSameSpecifiedColourInHtmlAndPdf(string name, string expectedHex) {
        using WordDocument source = WordDocument.Create();
        source.MainDocumentPartRoot.StyleDefinitionsPart!.Styles!.DocDefaults = new DocDefaults(
            new RunPropertiesDefault(new RunPropertiesBaseStyle(new Highlight { Val = new HighlightColorValues(name) })));
        source.AddParagraph("Highlight");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Contains("background-color:#" + expectedHex, document.ToHtml(), StringComparison.OrdinalIgnoreCase);
        using PdfPigDocument pdf = PdfPigDocument.Open((document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        })).ToBytes());
        double red = int.Parse(expectedHex.Substring(0, 2), NumberStyles.HexNumber) / 255D;
        double green = int.Parse(expectedHex.Substring(2, 2), NumberStyles.HexNumber) / 255D;
        double blue = int.Parse(expectedHex.Substring(4, 2), NumberStyles.HexNumber) / 255D;
        Assert.Contains(pdf.GetPage(1).Paths, path => {
            var (actualRed, actualGreen, actualBlue) = path.FillColor.ToRGBValues();
            // PDF colour operands are rounded; compare within half an 8-bit channel step.
            return path.IsFilled && Math.Abs(actualRed - red) < 0.5 / 255 && Math.Abs(actualGreen - green) < 0.5 / 255 && Math.Abs(actualBlue - blue) < 0.5 / 255;
        });
    }
    [Fact]
    public void ExactDarkMagentaImportUsesTheSharedPaletteWithoutApproximation() {
        HtmlToWordResult conversion = HtmlConversionDocument.Parse("<p><span style='background-color:#800080'>Highlight</span></p>")
            .ConvertToWordWithEvidence(new WordHtmlImportOptions { TextBackgroundMode = WordHtmlTextBackgroundMode.NearestHighlight });
        using WordDocument document = conversion.Value;
        Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "text-background-color-approximated");
        WordParagraph run = document.Paragraphs.SelectMany(paragraph => paragraph.GetRuns())
            .Single(paragraph => paragraph.Text == "Highlight");
        Assert.Equal(WordHighlightColor.DarkMagenta, run.Highlight);
        Assert.Contains("background-color:#800080", document.ToHtml(), StringComparison.OrdinalIgnoreCase);
    }

}
