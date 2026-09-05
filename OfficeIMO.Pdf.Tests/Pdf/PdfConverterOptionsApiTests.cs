using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConverterOptionsApiTests {
    [Fact]
    public void DirectAndComposedPdfOptionsUseUnambiguousStageNames() {
        Type html = typeof(OfficeIMO.Html.Pdf.HtmlToPdfOptions);
        Type asciiDoc = typeof(OfficeIMO.AsciiDoc.Pdf.AsciiDocToPdfOptions);
        Type latex = typeof(OfficeIMO.Latex.Pdf.LatexToPdfOptions);

        Assert.Equal(typeof(PdfOptions), html.GetProperty("PdfOptions")?.PropertyType);
        Assert.Null(html.GetProperty("DocumentOptions"));
        Assert.Equal(typeof(MarkdownToPdfOptions), asciiDoc.GetProperty("MarkdownOptions")?.PropertyType);
        Assert.Equal(typeof(MarkdownToPdfOptions), latex.GetProperty("MarkdownOptions")?.PropertyType);
        Assert.Null(asciiDoc.GetProperty("PdfOptions"));
        Assert.Null(latex.GetProperty("PdfOptions"));
    }

    [Fact]
    public void ConverterOptionsExposeExpectedResourceDefaults() {
        var markdown = new MarkdownToPdfOptions();
        var word = new OfficeIMO.Word.Pdf.WordToPdfOptions();
        var excel = new OfficeIMO.Excel.Pdf.ExcelToPdfOptions();
        var powerPoint = new OfficeIMO.PowerPoint.Pdf.PowerPointToPdfOptions();
        var html = new OfficeIMO.Html.Pdf.HtmlToPdfOptions();
        var rtf = new OfficeIMO.Rtf.Pdf.RtfToPdfOptions();
        var asciiDoc = new OfficeIMO.AsciiDoc.Pdf.AsciiDocToPdfOptions();
        var latex = new OfficeIMO.Latex.Pdf.LatexToPdfOptions();

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
        AssertBalancedDefault(asciiDoc.MarkdownOptions.ResourcePolicy);
        AssertBalancedDefault(latex.MarkdownOptions.ResourcePolicy);
        AssertPortable(PdfResourcePolicy.CreatePortableDeterministic());
    }

    [Fact]
    public void MarkdownProfileMapsToSingleOptionsObject() {
        var options = new MarkdownToPdfOptions {
            IncludeImages = true,
            ResourcePolicy = PdfResourcePolicy.CreateTrustedHost(),
            ApplyDefaultTheme = true
        };

        MarkdownToPdfOptions returned = options.UseProfile(PdfExportProfile.TextOnly);

        Assert.Same(options, returned);
        Assert.False(options.IncludeImages);
        Assert.True(options.ResourcePolicy.AllowDataUris);
        Assert.True(options.ResourcePolicy.AllowLocalFileAccess);
        Assert.True(options.ResourcePolicy.AllowDocumentFontEmbedding);
        Assert.False(options.ApplyDefaultTheme);
        Assert.Equal(MarkdownPdfFrontMatterRenderMode.Hidden, options.FrontMatterRenderMode);
    }

    [Fact]
    public void MarkdownProfilesDoNotDependOnPreviouslyAppliedProfile() {
        var reused = new MarkdownToPdfOptions();
        reused.UseProfile(PdfExportProfile.TextOnly).UseProfile(PdfExportProfile.Lightweight);
        var fresh = new MarkdownToPdfOptions().UseProfile(PdfExportProfile.Lightweight);

        Assert.Equal(fresh.IncludeImages, reused.IncludeImages);
        Assert.Equal(fresh.ApplyDefaultTheme, reused.ApplyDefaultTheme);
        Assert.Equal(fresh.CreateOutlineFromHeadings, reused.CreateOutlineFromHeadings);
        Assert.Equal(fresh.FrontMatterRenderMode, reused.FrontMatterRenderMode);
    }

    [Fact]
    public void WordProfileMapsPrintReadyChoices() {
        var options = new OfficeIMO.Word.Pdf.WordToPdfOptions {
            IncludePageNumbers = false,
            DefaultTableBorders = false
        };

        OfficeIMO.Word.Pdf.WordToPdfOptions returned = options.UseProfile(PdfExportProfile.PrintReady);

        Assert.Same(options, returned);
        Assert.False(options.IncludePageNumbers);
        Assert.True(options.DefaultTableBorders);
    }

    [Fact]
    public void ExcelProfileMapsLightweightChoices() {
        var options = new OfficeIMO.Excel.Pdf.ExcelToPdfOptions {
            UseWorksheetImages = true,
            UseWorksheetCharts = true,
            UseWorksheetHyperlinks = true
        };

        OfficeIMO.Excel.Pdf.ExcelToPdfOptions returned = options.UseProfile(PdfExportProfile.Lightweight);

        Assert.Same(options, returned);
        Assert.False(options.UseWorksheetHeaderFooterImages);
        Assert.False(options.UseWorksheetImages);
        Assert.False(options.UseWorksheetCharts);
        Assert.False(options.UseWorksheetHyperlinks);
        Assert.True(options.UseWorksheetCellStyles);
    }

    [Fact]
    public void ExcelProfilesDoNotDependOnPreviouslyAppliedProfile() {
        var reused = new OfficeIMO.Excel.Pdf.ExcelToPdfOptions();
        reused.UseProfile(PdfExportProfile.TextOnly).UseProfile(PdfExportProfile.PrintReady);
        var fresh = new OfficeIMO.Excel.Pdf.ExcelToPdfOptions().UseProfile(PdfExportProfile.PrintReady);

        Assert.Equal(GetExcelProfileState(fresh), GetExcelProfileState(reused));
    }

    [Fact]
    public void PowerPointProfileMapsTextOnlyChoices() {
        var options = new OfficeIMO.PowerPoint.Pdf.PowerPointToPdfOptions {
            IncludePictures = true,
            IncludeAutoShapes = true,
            IncludeCharts = true
        };

        OfficeIMO.PowerPoint.Pdf.PowerPointToPdfOptions returned = options.UseProfile(PdfExportProfile.TextOnly);

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

        Assert.Throws<ArgumentOutOfRangeException>(() => new MarkdownToPdfOptions().UseProfile(profile));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeIMO.Word.Pdf.WordToPdfOptions().UseProfile(profile));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeIMO.Excel.Pdf.ExcelToPdfOptions().UseProfile(profile));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeIMO.PowerPoint.Pdf.PowerPointToPdfOptions().UseProfile(profile));
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

    [Fact]
    public void PdfOptionsExposeAndClonePlannedFontFamilySubstitutions() {
        var options = new PdfOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Portable Latin", [1]))
            .RegisterFontFamilySubstitution(
                "Source Sans",
                "Portable Latin",
                PdfFontFamilySubstitutionImpact.Compatible);

        PdfOptions clone = options.Clone();
        PdfFontFamilySubstitution substitution = Assert.Single(options.FontFamilySubstitutions);
        PdfFontFamilySubstitution clonedSubstitution = Assert.Single(clone.FontFamilySubstitutions);

        Assert.Equal("Source Sans", substitution.SourceFontFamily);
        Assert.Equal("Portable Latin", substitution.TargetFontFamily);
        Assert.Equal(PdfFontFamilySubstitutionImpact.Compatible, substitution.Impact);
        Assert.True(options.HasNamedFontFamily("Source Sans"));
        Assert.True(clone.HasNamedFontFamily("Source Sans"));
        Assert.Equal(substitution.SourceFontFamily, clonedSubstitution.SourceFontFamily);
        Assert.Equal(substitution.TargetFontFamily, clonedSubstitution.TargetFontFamily);
        Assert.Equal(substitution.Impact, clonedSubstitution.Impact);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            options.RegisterFontFamilySubstitution(
                "Unknown impact source",
                "Portable Latin",
                (PdfFontFamilySubstitutionImpact)999));

        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Source Sans", [2]));
        Assert.True(options.TryResolveNamedFontFace("Source Sans", bold: false, italic: false, out PdfNamedFontFace face));
        Assert.Equal("Portable Latin", face.FamilyName);
    }

    [Fact]
    public void UnresolvedFontFamilySubstitutionRemainsAWarning() {
        var options = new PdfOptions()
            .RegisterFontFamilySubstitution(
                "Source Sans",
                "Missing Portable Latin",
                PdfFontFamilySubstitutionImpact.Compatible);

        PdfConversionWarning warning = options.CreateFontFamilySubstitutionWarning(
            "OfficeIMO.Tests",
            "FontFamilySubstituted",
            "document",
            "Source Sans",
            PdfStandardFont.Helvetica,
            resolvedFontFamily: null);

        Assert.Equal(PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.False(warning.Details.ContainsKey("plannedSubstitution"));
        Assert.DoesNotContain("Missing Portable Latin", warning.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ResolvedFallbackListSubstitutionUsesItsOwnImpactClassification() {
        const string resolvedFamily = "Portable Calibri";
        var options = new PdfOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(resolvedFamily, [1]))
            .RegisterFontFamilySubstitution(
                "Primary",
                "Missing Primary",
                PdfFontFamilySubstitutionImpact.LayoutSensitive)
            .RegisterFontFamilySubstitution(
                "Calibri",
                resolvedFamily,
                PdfFontFamilySubstitutionImpact.Compatible);

        Assert.True(options.TryResolveFontFamilySubstitution(
            "Primary, Calibri",
            out PdfFontFamilySubstitution? substitution));
        Assert.Equal("Calibri", substitution!.SourceFontFamily);

        PdfConversionWarning warning = options.CreateFontFamilySubstitutionWarning(
            "OfficeIMO.Tests",
            "FontFamilySubstituted",
            "document",
            "Primary, Calibri",
            fallbackSlot: null,
            resolvedFontFamily: resolvedFamily);

        Assert.Equal(PdfConversionWarningSeverity.Information, warning.Severity);
        Assert.Equal(bool.TrueString, warning.Details["plannedSubstitution"]);
        Assert.Equal(PdfFontFamilySubstitutionImpact.Compatible.ToString(), warning.Details["substitutionImpact"]);
        Assert.Contains(resolvedFamily, warning.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("Missing Primary", warning.Message, StringComparison.Ordinal);
    }

    private static void AssertPortable(PdfResourcePolicy policy) {
        Assert.False(policy.AllowSystemFontEmbedding);
        Assert.False(policy.AllowDocumentFontEmbedding);
        Assert.False(policy.AllowLocalFileAccess);
        Assert.False(policy.AllowRemoteResourceResolution);
        Assert.True(policy.AllowDataUris);
        Assert.True(policy.AllowEmbeddedPackageResources);
    }

    private static void AssertBalancedDefault(PdfResourcePolicy policy) {
        Assert.True(policy.AllowSystemFontEmbedding);
        Assert.False(policy.AllowDocumentFontEmbedding);
        Assert.False(policy.AllowLocalFileAccess);
        Assert.False(policy.AllowRemoteResourceResolution);
        Assert.True(policy.AllowDataUris);
        Assert.True(policy.AllowEmbeddedPackageResources);
    }

    private static bool[] GetExcelProfileState(OfficeIMO.Excel.Pdf.ExcelToPdfOptions options) => new[] {
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
