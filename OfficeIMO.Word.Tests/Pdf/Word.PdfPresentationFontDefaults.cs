using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "empty")]
    [InlineData(false, "header-size")]
    [InlineData(false, "footer-size")]
    [InlineData(false, "header-font")]
    [InlineData(false, "footer-font")]
    [InlineData(false, "header-family")]
    [InlineData(false, "footer-family")]
    [InlineData(false, "metadata")]
    [InlineData(true, "empty")]
    [InlineData(true, "header-size")]
    [InlineData(true, "footer-size")]
    [InlineData(true, "header-font")]
    [InlineData(true, "footer-font")]
    [InlineData(true, "header-family")]
    [InlineData(true, "footer-family")]
    [InlineData(true, "metadata")]
    public void SaveAsPdf_PresentationOptionsPreserveInheritedBodyTypography(bool profile, string setting) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Times New Roman";
        document.Settings.FontSize = 18;
        document.AddParagraph("Inherited body font and size");
        var options = new WordToPdfOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        if (profile) options.UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));
        else options.PdfOptions = new PdfOptions();
        PdfOptions configured = options.PdfOptions!;
        switch (setting) {
            case "header-size": configured.HeaderFontSize = 24; break;
            case "footer-size": configured.FooterFontSize = 24; break;
            case "header-font": configured.HeaderFont = PdfStandardFont.Courier; break;
            case "footer-font": configured.FooterFont = PdfStandardFont.Courier; break;
            case "header-family": configured.HeaderFontFamily = "Courier New"; break;
            case "footer-family": configured.FooterFontFamily = "Courier New"; break;
            case "metadata": options.Title = "Typography control"; break;
        }
        string source = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var baseline = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        using var actual = PdfPigDocument.Open(document.ToPdfBytes(options));
        var expected = baseline.GetPage(1).Letters;
        var letters = actual.GetPage(1).Letters;
        Assert.All(letters, letter => Assert.Contains("Times", letter.FontName, StringComparison.Ordinal));
        Assert.All(letters, letter => Assert.Equal(18D, letter.PointSize, 3));
        Assert.Equal(expected.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.StartBaseLine)),
            letters.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.StartBaseLine)));
        Assert.Equal(source, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ExplicitBodyFontSurvivesPresentationOverrides(bool profile) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Times New Roman";
        document.AddParagraph("Body font override");
        var options = new WordToPdfOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        if (profile) options.UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));
        else options.PdfOptions = new PdfOptions();
        options.PdfOptions!.DefaultFont = PdfStandardFont.Courier;
        options.PdfOptions.HeaderFontSize = 24;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(options));
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Contains("Courier", letter.FontName, StringComparison.Ordinal));
    }
}
