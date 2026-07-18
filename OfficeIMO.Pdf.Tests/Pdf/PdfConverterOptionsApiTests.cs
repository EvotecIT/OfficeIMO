using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConverterOptionsApiTests {
    [Fact]
    public void ConverterOptionsExposeUnifiedBalancedResourceDefaults() {
        var markdown = new MarkdownPdfSaveOptions();
        var word = new OfficeIMO.Word.Pdf.PdfSaveOptions();
        var excel = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions();
        var powerPoint = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions();
        var html = new OfficeIMO.Html.Pdf.HtmlPdfSaveOptions();
        var rtf = new OfficeIMO.Rtf.Pdf.RtfPdfSaveOptions();
        var asciiDoc = new OfficeIMO.AsciiDoc.Pdf.AsciiDocPdfSaveOptions();
        var latex = new OfficeIMO.Latex.Pdf.LatexPdfSaveOptions();

        Assert.Equal(PdfTextFallbackFeatures.Default, markdown.TextFallbacks);
        Assert.Equal(PdfTextFallbackFeatures.Default, word.TextFallbacks);
        Assert.Equal(PdfTextFallbackFeatures.Default, excel.TextFallbacks);
        Assert.Equal(PdfTextFallbackFeatures.Default, powerPoint.TextFallbacks);
        AssertBalancedDefault(markdown.ResourcePolicy);
        AssertBalancedDefault(word.ResourcePolicy);
        AssertBalancedDefault(excel.ResourcePolicy);
        AssertBalancedDefault(powerPoint.ResourcePolicy);
        AssertBalancedDefault(html.ResourcePolicy);
        AssertBalancedDefault(rtf.ResourcePolicy);
        AssertBalancedDefault(asciiDoc.PdfOptions.ResourcePolicy);
        AssertBalancedDefault(latex.PdfOptions.ResourcePolicy);
        AssertPortable(PdfResourcePolicy.CreatePortableDeterministic());
    }

    [Fact]
    public void MarkdownProfileMapsToSingleOptionsObject() {
        var options = new MarkdownPdfSaveOptions {
            IncludeImages = true,
            ResourcePolicy = PdfResourcePolicy.CreateTrustedHost(),
            ApplyDefaultTheme = true
        };

        MarkdownPdfSaveOptions returned = options.UseProfile(PdfExportProfile.TextOnly);

        Assert.Same(options, returned);
        Assert.False(options.IncludeImages);
        Assert.True(options.ResourcePolicy.AllowDataUris);
        Assert.True(options.ResourcePolicy.AllowLocalFileAccess);
        Assert.False(options.ApplyDefaultTheme);
        Assert.Equal(MarkdownPdfFrontMatterRenderMode.Hidden, options.FrontMatterRenderMode);
    }

    [Fact]
    public void MarkdownProfilesDoNotDependOnPreviouslyAppliedProfile() {
        var reused = new MarkdownPdfSaveOptions();
        reused.UseProfile(PdfExportProfile.TextOnly).UseProfile(PdfExportProfile.Lightweight);
        var fresh = new MarkdownPdfSaveOptions().UseProfile(PdfExportProfile.Lightweight);

        Assert.Equal(fresh.IncludeImages, reused.IncludeImages);
        Assert.Equal(fresh.ApplyDefaultTheme, reused.ApplyDefaultTheme);
        Assert.Equal(fresh.CreateOutlineFromHeadings, reused.CreateOutlineFromHeadings);
        Assert.Equal(fresh.FrontMatterRenderMode, reused.FrontMatterRenderMode);
    }

    [Fact]
    public void WordProfileMapsPrintReadyChoices() {
        var options = new OfficeIMO.Word.Pdf.PdfSaveOptions {
            IncludePageNumbers = false,
            DefaultTableBorders = false
        };

        OfficeIMO.Word.Pdf.PdfSaveOptions returned = options.UseProfile(PdfExportProfile.PrintReady);

        Assert.Same(options, returned);
        Assert.False(options.IncludePageNumbers);
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
    public void ExcelProfilesDoNotDependOnPreviouslyAppliedProfile() {
        var reused = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions();
        reused.UseProfile(PdfExportProfile.TextOnly).UseProfile(PdfExportProfile.PrintReady);
        var fresh = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions().UseProfile(PdfExportProfile.PrintReady);

        Assert.Equal(GetExcelProfileState(fresh), GetExcelProfileState(reused));
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

    [Fact]
    public void PdfOptionsTrackAndCloneExplicitFontSlots() {
        var defaults = new PdfOptions();
        PdfOptions defaultClone = defaults.Clone();

        Assert.False(defaults.HasExplicitDefaultFont);
        Assert.False(defaults.HasExplicitHeaderFont);
        Assert.False(defaults.HasExplicitFooterFont);
        Assert.False(defaultClone.HasExplicitDefaultFont);
        Assert.False(defaultClone.HasExplicitHeaderFont);
        Assert.False(defaultClone.HasExplicitFooterFont);

        var configured = new PdfOptions {
            DefaultFont = PdfStandardFont.Courier,
            HeaderFont = PdfStandardFont.TimesRoman,
            FooterFont = PdfStandardFont.Helvetica
        };
        PdfOptions configuredClone = configured.Clone();

        Assert.True(configured.HasExplicitDefaultFont);
        Assert.True(configured.HasExplicitHeaderFont);
        Assert.True(configured.HasExplicitFooterFont);
        Assert.True(configuredClone.HasExplicitDefaultFont);
        Assert.True(configuredClone.HasExplicitHeaderFont);
        Assert.True(configuredClone.HasExplicitFooterFont);
    }

    private static void AssertPortable(PdfResourcePolicy policy) {
        Assert.False(policy.AllowSystemFontEmbedding);
        Assert.False(policy.AllowLocalFileAccess);
        Assert.False(policy.AllowRemoteResourceResolution);
        Assert.True(policy.AllowDataUris);
        Assert.True(policy.AllowEmbeddedPackageResources);
    }

    private static void AssertBalancedDefault(PdfResourcePolicy policy) {
        Assert.True(policy.AllowSystemFontEmbedding);
        Assert.False(policy.AllowLocalFileAccess);
        Assert.False(policy.AllowRemoteResourceResolution);
        Assert.True(policy.AllowDataUris);
        Assert.True(policy.AllowEmbeddedPackageResources);
    }

    private static bool[] GetExcelProfileState(OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions options) => new[] {
        options.RespectWorkbookSheetVisibility,
        options.UseWorksheetPrintAreas,
        options.UseWorksheetPageSetup,
        options.UseWorksheetPrintTitleRows,
        options.UseWorksheetPageBreaks,
        options.UseWorksheetHeadersAndFooters,
        options.UseWorksheetHeaderFooterImages,
        options.UseWorksheetCellStyles,
        options.UseWorksheetHyperlinks,
        options.UseWorksheetImages,
        options.UseWorksheetCharts,
        options.UseWorksheetMergedCells,
        options.UseWorksheetColumnWidths,
        options.UseWorksheetRowHeights,
        options.RespectWorksheetHiddenRowsAndColumns,
        options.IncludeSheetHeadings
    };
}
