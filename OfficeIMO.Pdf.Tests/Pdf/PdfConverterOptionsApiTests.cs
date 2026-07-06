using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConverterOptionsApiTests {
    [Fact]
    public void ConverterOptionsExposeUnifiedTextFallbackDefaults() {
        var markdown = new MarkdownPdfSaveOptions();
        var word = new OfficeIMO.Word.Pdf.PdfSaveOptions();
        var excel = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions();
        var powerPoint = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions();

        Assert.Equal(PdfTextFallbackFeatures.Default, markdown.TextFallbacks);
        Assert.Equal(PdfTextFallbackFeatures.Default, word.TextFallbacks);
        Assert.Equal(PdfTextFallbackFeatures.Default, excel.TextFallbacks);
        Assert.Equal(PdfTextFallbackFeatures.Default, powerPoint.TextFallbacks);
        Assert.True(markdown.AllowSystemFontEmbedding);
        Assert.False(word.AllowSystemFontEmbedding);
        Assert.True(excel.AllowSystemFontEmbedding);
        Assert.True(powerPoint.AllowSystemFontEmbedding);
    }

    [Fact]
    public void MarkdownProfileMapsToSingleOptionsObject() {
        var options = new MarkdownPdfSaveOptions {
            IncludeDataUriImages = true,
            IncludeLocalImages = true,
            ApplyWordLikeTheme = true
        };

        MarkdownPdfSaveOptions returned = options.UseProfile(PdfExportProfile.TextOnly);

        Assert.Same(options, returned);
        Assert.False(options.IncludeDataUriImages);
        Assert.False(options.IncludeLocalImages);
        Assert.False(options.ApplyWordLikeTheme);
        Assert.Equal(MarkdownPdfFrontMatterRenderMode.Hidden, options.FrontMatterRenderMode);
    }

    [Fact]
    public void WordProfileMapsPrintReadyChoices() {
        var options = new OfficeIMO.Word.Pdf.PdfSaveOptions {
            IncludePageNumbers = false,
            DefaultTableBorders = false
        };

        OfficeIMO.Word.Pdf.PdfSaveOptions returned = options.UseProfile(PdfExportProfile.PrintReady);

        Assert.Same(options, returned);
        Assert.True(options.IncludePageNumbers);
        Assert.True(options.DefaultTableBorders);
    }

    [Fact]
    public void ExcelProfileMapsLightweightChoices() {
        var options = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions {
            UseWorksheetImages = true,
            UseWorksheetCharts = true,
            UseWorksheetHyperlinks = true
        };

        OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions returned = options.UseProfile(PdfExportProfile.Lightweight);

        Assert.Same(options, returned);
        Assert.False(options.UseWorksheetHeaderFooterImages);
        Assert.False(options.UseWorksheetImages);
        Assert.False(options.UseWorksheetCharts);
        Assert.False(options.UseWorksheetHyperlinks);
        Assert.True(options.UseWorksheetCellStyles);
    }

    [Fact]
    public void PowerPointProfileMapsTextOnlyChoices() {
        var options = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions {
            IncludePictures = true,
            IncludeAutoShapes = true,
            IncludeCharts = true
        };

        OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions returned = options.UseProfile(PdfExportProfile.TextOnly);

        Assert.Same(options, returned);
        Assert.False(options.IncludePictures);
        Assert.False(options.IncludeAutoShapes);
        Assert.True(options.IncludeTextBoxes);
        Assert.True(options.IncludeTables);
        Assert.False(options.IncludeCharts);
    }

    [Fact]
    public void ConverterProfilesRejectUnknownValues() {
        var profile = (PdfExportProfile)999;

        Assert.Throws<ArgumentOutOfRangeException>(() => new MarkdownPdfSaveOptions().UseProfile(profile));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeIMO.Word.Pdf.PdfSaveOptions().UseProfile(profile));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions().UseProfile(profile));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions().UseProfile(profile));
    }
}
