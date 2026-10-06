using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(24480U, 15840U)]
    [InlineData(12240U, 10000U)]
    public void PageSize_NativeLandscapePreservesElidedDefaultDimensions(uint width, uint height) {
        using WordDocument document = WordDocument.Create();
        WordPageSizes settings = document.Sections[0].PageSettings;
        settings.Orientation = OfficePageOrientation.Landscape;
        settings.Width = width; settings.Height = height;
        document.AddParagraph("Landscape default dimension");
        using var source = new MemoryStream(document.ToBytes(WordFileFormat.Doc));
        using WordDocument loaded = WordDocument.Load(source);
        Assert.Equal(width, loaded.Sections[0].PageSettings.Width);
        Assert.Equal(height, loaded.Sections[0].PageSettings.Height);
        Assert.Equal(OfficePageOrientation.Landscape, loaded.Sections[0].PageSettings.Orientation);
        PdfCore.PdfPageInfo page = Assert.Single(PdfCore.PdfInspector.Inspect(loaded.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false
        })).Pages);
        Assert.Equal(width / 20D, page.Width);
        Assert.Equal(height / 20D, page.Height);
    }

    public static IEnumerable<object[]> AdditionalWordPaperPresets() {
        object[][] presets = {
            new object[] { WordPageSize.Tabloid, 15840U, 24480U, (ushort)3 },
            new object[] { WordPageSize.B4Jis, 14570U, 20636U, (ushort)12 },
            new object[] { WordPageSize.Envelope9, 5580U, 12780U, (ushort)19 },
            new object[] { WordPageSize.Envelope10, 5940U, 13680U, (ushort)20 },
            new object[] { WordPageSize.CSheet, 24480U, 31680U, (ushort)24 },
            new object[] { WordPageSize.EnvelopeDl, 6236U, 12472U, (ushort)27 },
            new object[] { WordPageSize.EnvelopeC5, 9184U, 12983U, (ushort)28 },
            new object[] { WordPageSize.EnvelopeC4, 12983U, 18369U, (ushort)30 },
            new object[] { WordPageSize.EnvelopeB5, 9978U, 14173U, (ushort)34 },
            new object[] { WordPageSize.EnvelopeMonarch, 5580U, 10800U, (ushort)37 }
        };
        foreach (object[] preset in presets) {
            foreach (bool landscape in new[] { false, true }) {
                foreach (WordFileFormat format in new[] { WordFileFormat.Docx, WordFileFormat.Doc }) {
                    yield return preset.Concat(new object[] { landscape, format }).ToArray();
                }
            }
        }
    }

    [Theory]
    [MemberData(nameof(AdditionalWordPaperPresets))]
    public void PageSize_AdditionalPresetsSurviveNativeFormatsAndPdf(WordPageSize preset, uint width, uint height,
        ushort paperCode, bool landscape, WordFileFormat format) {
        using WordDocument document = WordDocument.Create();
        WordPageSizes settings = document.Sections[0].PageSettings;
        settings.Orientation = landscape ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
        settings.PageSize = preset;
        document.AddParagraph("Preset marker");
        Assert.Equal(preset, settings.PageSize);
        Assert.Equal(paperCode, settings.Code);
        WordPageSizeDefinition definition = WordPageSizes.GetDefinition(preset)!;
        Assert.Equal(width, definition.WidthTwips);
        Assert.Equal(height, definition.HeightTwips);
        Assert.Equal(paperCode, definition.PaperCode);
        Assert.Empty(new OpenXmlValidator().Validate(document.Sections[0]._sectionProperties));
        using var source = new MemoryStream(document.ToBytes(format));
        using WordDocument loaded = WordDocument.Load(source);
        WordPageSizes reopened = loaded.Sections[0].PageSettings;
        Assert.Empty(new OpenXmlValidator().Validate(loaded.Sections[0]._sectionProperties));
        Assert.Equal(preset, reopened.PageSize);
        Assert.Equal(landscape ? height : width, reopened.Width);
        Assert.Equal(landscape ? width : height, reopened.Height);
        Assert.Equal(settings.Orientation, reopened.Orientation);
        Assert.Equal("Preset marker", loaded.Paragraphs[0].Text);
        PdfCore.PdfPageInfo page = Assert.Single(PdfCore.PdfInspector.Inspect(loaded.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false
        })).Pages);
        Assert.Equal((landscape ? height : width) / 20D, page.Width, 3);
        Assert.Equal((landscape ? width : height) / 20D, page.Height, 3);
    }

    [Theory]
    [InlineData(WordPageSize.Envelope10, 297D, 684D)]
    [InlineData(WordPageSize.B5, 515.9D, 728.5D)]
    public void SaveAsPdf_DefaultPaperPresetUsesTheWordDimensionCatalog(WordPageSize preset, double width, double height) {
        using WordDocument document = WordDocument.Create();
        document.Sections[0]._sectionProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>()!.Remove();
        document.AddParagraph("Default preset");
        PdfCore.PdfPageInfo page = Assert.Single(PdfCore.PdfInspector.Inspect(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, DefaultPageSize = preset
        })).Pages);
        Assert.Equal(width, page.Width, 3);
        Assert.Equal(height, page.Height, 3);
    }

    [Theory]
    [InlineData(0U)]
    [InlineData(1U)]
    public void PageSize_RecognizesProducerDimensionsWithoutPrinterCode(uint roundingDifference) {
        using WordDocument document = WordDocument.Create();
        WordPageSizes settings = document.Sections[0].PageSettings;
        settings.PageSize = WordPageSize.A4;
        settings.Code = null;
        settings.Width = WordPageSizes.A4.WidthTwips + roundingDifference;
        settings.Height = WordPageSizes.A4.HeightTwips + roundingDifference;
        Assert.Equal(WordPageSize.A4, settings.PageSize);
    }

    [Fact]
    public void PageSize_ReplacementRetainsSectionSchemaOrder() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        var validator = new OpenXmlValidator();
        Assert.Empty(validator.Validate(section._sectionProperties));
        section.PageSettings.PageSize = WordPageSize.Legal;
        Assert.Empty(validator.Validate(section._sectionProperties));
    }

    [Fact]
    public void SaveAsPdf_AuthoredWideDimensionsWithoutOrientationRemainWide() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        section.PageSettings.Width = 14000;
        section.PageSettings.Height = 8000;
        section._sectionProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>()!.Orient = null;
        document.AddParagraph("Wide page");
        PdfCore.PdfPageInfo page = Assert.Single(PdfCore.PdfInspector.Inspect(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false
        })).Pages);
        Assert.Equal(700D, page.Width);
        Assert.Equal(400D, page.Height);
    }
}
