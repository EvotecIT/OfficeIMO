using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void SaveAsPdf_NamedPresentationFamilyPreservesBodyTypography(bool profile, bool footer, bool explicitBody) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Times New Roman";
        document.Settings.FontSize = 18;
        document.AddParagraph("BODY");
        WordPdfExportOptions options = CreateNamedPresentationOptions(profile);
        if (explicitBody) options.PdfOptions!.DefaultFont = PdfStandardFont.Courier;
        if (footer) {
            options.PdfOptions!.FooterFontFamily = ManagedTextShapingTestAssets.FamilyName;
            options.IncludePageNumbers = true;
            options.PageNumberFormat = "NUMBER {page}/{pages}";
        } else {
            options.PdfOptions!.HeaderFontFamily = ManagedTextShapingTestAssets.FamilyName;
            options.PdfOptions.ShowHeader = true;
            options.PdfOptions.HeaderFormat = "HEADER";
        }
        var baselineOptions = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        if (explicitBody) baselineOptions.PdfOptions = new PdfCreateOptions { DefaultFont = PdfStandardFont.Courier };
        string source = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var baseline = PdfPigDocument.Open(document.ToPdf(baselineOptions).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        using var actual = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var expected = Assert.Single(baseline.GetPage(1).GetWords(), word => word.Text == "BODY").Letters;
        var body = Assert.Single(actual.GetPage(1).GetWords(), word => word.Text == "BODY").Letters;
        Assert.All(body, letter => {
            Assert.Contains(explicitBody ? "Courier" : "Times", letter.FontName, StringComparison.Ordinal);
            Assert.Equal(18D, letter.PointSize, 3);
        });
        Assert.Equal(expected.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.StartBaseLine)),
            body.Select(letter => (letter.Value, letter.FontName, letter.PointSize, letter.StartBaseLine)));
        var presentation = actual.GetPage(1).Letters.Where(letter => letter.PointSize == 9).ToArray();
        Assert.Equal(footer ? "NUMBER1/1" : "HEADER", string.Concat(presentation.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).Select(letter => letter.Value)));
        Assert.All(presentation, letter => Assert.Contains("OfficeIMO", letter.FontName, StringComparison.Ordinal));
        Assert.Equal(source, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_UnusedNamedFamilyPreservesBodyTypography(bool profile) {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Times New Roman";
        document.Settings.FontSize = 18;
        document.AddParagraph("BODY");
        using var pdf = PdfPigDocument.Open(document.ToPdf(CreateNamedPresentationOptions(profile)).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.All(pdf.GetPage(1).Letters, letter => {
            Assert.Contains("Times", letter.FontName, StringComparison.Ordinal);
            Assert.Equal(18D, letter.PointSize, 3);
        });
    }

    private static WordPdfExportOptions CreateNamedPresentationOptions(bool profile) {
        var options = new WordPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        if (profile) options.UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));
        else options.PdfOptions = new PdfCreateOptions();
        options.PdfOptions!.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        return options;
    }
}
