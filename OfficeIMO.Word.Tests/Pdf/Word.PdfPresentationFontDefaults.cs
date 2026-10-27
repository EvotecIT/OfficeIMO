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
        var options = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        if (profile) options.UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));
        else options.PdfOptions = new PdfCreateOptions();
        PdfCreateOptions configured = options.PdfOptions!;
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
        using var baseline = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        using var actual = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        var options = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        if (profile) options.UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));
        else options.PdfOptions = new PdfCreateOptions();
        options.PdfOptions!.DefaultFont = PdfStandardFont.Courier;
        options.PdfOptions.HeaderFontSize = 24;
        using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Contains("Courier", letter.FontName, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_GeneratedPresentationRetainsItsExplicitFont(bool profile, bool footer) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Times New Roman";
        document.Settings.FontSize = 18;
        document.AddParagraph("BODY");
        var options = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        if (profile) options.UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));
        else options.PdfOptions = new PdfCreateOptions();
        if (footer) {
            options.PdfOptions!.FooterFont = PdfStandardFont.Courier;
            options.IncludePageNumbers = true;
            options.PageNumberFormat = "NUMBER {page}/{pages}";
        } else {
            options.PdfOptions!.HeaderFont = PdfStandardFont.Courier;
            options.PdfOptions.ShowHeader = true;
            options.PdfOptions.HeaderFormat = "HEADER";
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(footer ? "BODYNUMBER1/1" : "HEADERBODY", string.Concat(letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).Select(letter => letter.Value)));
        Assert.All(letters.Where(letter => letter.PointSize == 18), letter => Assert.Contains("Times", letter.FontName, StringComparison.Ordinal));
        Assert.All(letters.Where(letter => letter.PointSize == 9), letter => Assert.Contains("Courier", letter.FontName, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_DefaultSizeAssignmentPreservesTheInheritedFamily(bool profile, bool repeatProfile) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Times New Roman";
        document.Settings.FontSize = 18;
        document.AddParagraph("BODY");
        var options = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        var fontless = new OfficeRenderingProfile("shaping-only");
        if (profile) options.UseRenderingProfile(fontless);
        else options.PdfOptions = new PdfCreateOptions();
        double size = repeatProfile ? 13 : options.PdfOptions!.DefaultFontSize;
        options.PdfOptions!.DefaultFontSize = size;
        if (repeatProfile) options.UseRenderingProfile(fontless, OfficeRenderingProfileApplyMode.Replace);
        using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.All(pdf.GetPage(1).Letters, letter => {
            Assert.Equal(size, letter.PointSize, 3);
            Assert.Contains("Times", letter.FontName, StringComparison.Ordinal);
        });
    }
}
