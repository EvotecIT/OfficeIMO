using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfRenderingProfileTests {
    [Fact]
    public void SharedRenderingProfileConfiguresGeneratedPdfText() {
        var profile = new OfficeRenderingProfile(
            "managed-arabic",
            textShapingProvider: OfficeManagedTextShapingProvider.Instance,
            textShapingLanguage: " ar ");
        var options = new PdfOptions();

        PdfOptions returned = options.UseRenderingProfile(profile);

        Assert.Same(options, returned);
        Assert.Same(OfficeManagedTextShapingProvider.Instance, options.TextShapingProvider);
        Assert.Equal("ar", options.Language);
    }

    [Fact]
    public void OverlayPreservesExistingProviderWhenProfileDeclinesToOwnIt() {
        var existing = new DecliningTextShapingProvider();
        var options = new PdfOptions {
            TextShapingProvider = existing,
            Language = "pl"
        };

        options.UseRenderingProfile(
            new OfficeRenderingProfile("fonts-only"),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.Same(existing, options.TextShapingProvider);
        Assert.Equal("pl", options.Language);
    }

    [Fact]
    public void SharedRenderingProfileRegistersDeterministicFontsAndFallbacks() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Profile Sans", ManagedTextShapingTestAssets.CreateFont(' ', 'A'))
            .AddFallbackFamily("Profile Sans");
        var options = new PdfOptions();

        options.UseRenderingProfile(new OfficeRenderingProfile("portable", fonts));

        Assert.True(options.HasNamedFontFamily("Profile Sans"));
        Assert.Equal(
            new[] { "Profile Sans" },
            options.EmbeddedFontFallbacks?.FontFamilyNames);
    }

    [Fact]
    public void SharedRenderingProfileDoesNotPromoteNamedFontsIntoUndeclaredFallbacks() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Named Only", ManagedTextShapingTestAssets.CreateFont('A'));
        var options = new PdfOptions();

        options.UseRenderingProfile(new OfficeRenderingProfile("named-only", fonts));

        Assert.True(options.HasNamedFontFamily("Named Only"));
        Assert.Null(options.EmbeddedFontFallbacks);
    }

    [Fact]
    public void SharedRenderingProfilePreservesFallbackOrderAndUnicodeRanges() {
        var onlyA = new OfficeFontUnicodeRangeSet(new[] {
            new OfficeFontUnicodeRange('A', 'A')
        });
        var fonts = new OfficeFontFaceCollection()
            .Add(
                "First",
                ManagedTextShapingTestAssets.CreateFont('A', 'B'),
                OfficeFontStyle.Regular,
                onlyA)
            .Add("Second", ManagedTextShapingTestAssets.CreateFont('A', 'B'))
            .AddFallbackFamily("First")
            .AddFallbackFamily("Second");
        var options = new PdfOptions();

        options.UseRenderingProfile(new OfficeRenderingProfile("ranged", fonts));

        PdfEmbeddedFontFallbackSet fallbacks = Assert.IsType<PdfEmbeddedFontFallbackSet>(
            options.EmbeddedFontFallbacks);
        Assert.Equal(
            new[] {
                fonts.Faces[0].ResourceFamilyName,
                fonts.Faces[1].ResourceFamilyName
            },
            fallbacks.FontFamilyNames);
        PdfTextFallbackSegment segment = Assert.Single(fallbacks.PlanText("B").Segments);
        Assert.Equal(1, segment.FontIndex);
        Assert.Equal(fonts.Faces[1].ResourceFamilyName, segment.FontName);
    }

    [Fact]
    public void OverlayPreservesExplicitPdfFallbacksWhileRegisteringProfileFonts() {
        var options = new PdfOptions()
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate(
                        "Existing Fallback",
                        ManagedTextShapingTestAssets.CreateFont('B'))
                }));
        var profileFonts = new OfficeFontFaceCollection()
            .Add("Profile Sans", ManagedTextShapingTestAssets.CreateFont('A'));

        options.UseRenderingProfile(
            new OfficeRenderingProfile("overlay", profileFonts),
            OfficeRenderingProfileApplyMode.Overlay);

        Assert.True(options.HasNamedFontFamily("Profile Sans"));
        Assert.Equal(
            new[] { "Existing Fallback" },
            options.EmbeddedFontFallbacks?.FontFamilyNames);
    }

    [Fact]
    public void ReplaceClearsPreviouslyRegisteredPdfFontState() {
        var options = new PdfOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                "Existing Family",
                ManagedTextShapingTestAssets.CreateFont('A')))
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(
                new[] {
                    new PdfEmbeddedFontFallbackCandidate(
                        "Existing Fallback",
                        ManagedTextShapingTestAssets.CreateFont('B'))
                }));

        options.UseRenderingProfile(new OfficeRenderingProfile("managed-only"));

        Assert.Empty(options.NamedFontFamilies);
        Assert.Null(options.EmbeddedFontFallbacks);
    }

    [Fact]
    public void SharedRenderingProfileFlowsThroughFirstPartyOfficePdfAdapters() {
        var profile = new OfficeRenderingProfile(
            "managed-polish",
            textShapingProvider: OfficeManagedTextShapingProvider.Instance,
            textShapingLanguage: "pl");

        var word = new OfficeIMO.Word.Pdf.PdfSaveOptions().UseRenderingProfile(profile);
        var excel = new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions().UseRenderingProfile(profile);
        var powerPoint = new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions()
            .UseRenderingProfile(profile);

        Assert.Same(OfficeManagedTextShapingProvider.Instance, word.PdfOptions?.TextShapingProvider);
        Assert.Same(OfficeManagedTextShapingProvider.Instance, excel.PdfOptions?.TextShapingProvider);
        Assert.Same(OfficeManagedTextShapingProvider.Instance, powerPoint.PdfOptions?.TextShapingProvider);
        Assert.Equal("pl", word.PdfOptions?.Language);
        Assert.Equal("pl", excel.PdfOptions?.Language);
        Assert.Equal("pl", powerPoint.PdfOptions?.Language);
    }

    [Fact]
    public void SharedRenderingProfileSurvivesOfficeAdapterCloningAndPdfGeneration() {
        OfficeRenderingProfile profile = OfficeRenderingProfile.Managed;
        byte[] wordPdf;
        using (var wordStream = new MemoryStream())
        using (OfficeIMO.Word.WordDocument word = OfficeIMO.Word.WordDocument.Create(wordStream)) {
            word.AddParagraph("Word profile proof");
            wordPdf = OfficeIMO.Word.Pdf.WordPdfConverterExtensions.ToPdf(
                word,
                new OfficeIMO.Word.Pdf.PdfSaveOptions().UseRenderingProfile(profile));
        }

        byte[] excelPdf;
        using (OfficeIMO.Excel.ExcelDocument excel =
            OfficeIMO.Excel.ExcelDocument.Create(new MemoryStream())) {
            excel.AddWorksheet("Profile").CellValue(1, 1, "Excel profile proof");
            excelPdf = OfficeIMO.Excel.Pdf.ExcelPdfConverterExtensions.ToPdf(
                excel,
                new OfficeIMO.Excel.Pdf.ExcelPdfSaveOptions().UseRenderingProfile(profile));
        }

        byte[] powerPointPdf;
        using (OfficeIMO.PowerPoint.PowerPointPresentation powerPoint =
            OfficeIMO.PowerPoint.PowerPointPresentation.Create(new MemoryStream())) {
            powerPoint.AddSlide().AddTextBoxPoints(
                "PowerPoint profile proof",
                24,
                24,
                240,
                40);
            powerPointPdf = OfficeIMO.PowerPoint.Pdf.PowerPointPdfConverterExtensions.ToPdf(
                powerPoint,
                new OfficeIMO.PowerPoint.Pdf.PowerPointPdfSaveOptions()
                    .UseRenderingProfile(profile));
        }

        Assert.Contains("Word profile proof", PdfReadDocument.Open(wordPdf).ExtractText());
        Assert.Contains("Excel profile proof", PdfReadDocument.Open(excelPdf).ExtractText());
        Assert.Contains("PowerPoint profile proof", PdfReadDocument.Open(powerPointPdf).ExtractText());
    }

    private sealed class DecliningTextShapingProvider : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => null;
    }
}
