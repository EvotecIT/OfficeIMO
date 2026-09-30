using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordPdfOmittedFontSizeTests {
    // Reduced templates reproduce Word's fallbacks without importing private documents.
    [Theory]
    [InlineData(false, null, 12D)]
    [InlineData(true, null, 10D)]
    [InlineData(true, "22", 11D)]
    [InlineData(true, "21", 10.5D)]
    public void MissingSizeUsesWordFallbackWhileDeclaredDefaultsRemainExact(bool keepDefaults, string? size, double expectedSize) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        foreach (FontSize fontSize in styles.Descendants<FontSize>().ToArray()) fontSize.Remove();
        foreach (FontSizeComplexScript fontSize in styles.Descendants<FontSizeComplexScript>().ToArray()) fontSize.Remove();
        if (!keepDefaults) styles.DocDefaults?.Remove();
        if (size != null) styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new FontSize { Val = size });

        document.AddParagraph("Body");
        document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Cell";
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
        Assert.Equal("BodyCell", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => Assert.Equal(expectedSize, letter.PointSize, precision: 3));
    }

    [Fact]
    public void PolicySubstitutionReportsThePolicyRatherThanAMissingFont() {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Arial";
        document.AddParagraph("Policy test");
        using var stream = new System.IO.MemoryStream();
        var report = document.SaveAsPdf(stream, new WordToPdfOptions {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            IncludePageNumbers = false
        });
        report.RequireSuccess();
        var warning = Assert.Single(report.Warnings, warning =>
            warning.Code == "NativeFontFamilySubstituted" && warning.Details["fontFamily"] == "Arial");
        Assert.Equal("ResourcePolicy", warning.Details["substitutionReason"]);
        Assert.Contains("resource policy", warning.Message, System.StringComparison.Ordinal);
        Assert.DoesNotContain("unavailable", warning.Message, System.StringComparison.Ordinal);
    }
}
