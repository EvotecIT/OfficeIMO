using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfListMarkerKerningTests {
    [Theory]
    [InlineData("body", 0U, 10D)]
    [InlineData("cell", 0U, 10D)]
    [InlineData("header", 0U, 10D)]
    [InlineData("footer", 0U, 10D)]
    [InlineData("body", 24U, 9.2D)]
    [InlineData("cell", 24U, 9.2D)]
    [InlineData("header", 24U, 9.2D)]
    [InlineData("footer", 24U, 9.2D)]
    public void NumberingLevelKerningOverridesReachMarkerOutput(string route, uint threshold, double advance) =>
        AssertMarkerAdvance(route, defaultThreshold: threshold == 0 ? 2U : 0U,
            markerThreshold: threshold, paragraphMarkThreshold: null, fullOverride: false, advance);

    [Fact]
    public void InlineMarkerUsesItsOwnFontSizeForInheritedKerning() =>
        AssertMarkerAdvance("body", 24U, null, null, false, 9.2D);

    [Theory]
    [InlineData(0U, 10D)]
    [InlineData(2U, 9.2D)]
    public void ParagraphMarkKerningOverridesBodyFormatting(uint threshold, double advance) =>
        AssertMarkerAdvance("body", threshold == 0 ? 2U : 0U, null, threshold, false, advance);

    [Theory]
    [InlineData(0U, 10D)]
    [InlineData(24U, 9.2D)]
    public void FullNumberingOverrideReplacesAbstractLevelKerning(uint threshold, double advance) =>
        AssertMarkerAdvance("body", threshold == 0 ? 2U : 0U, threshold, null, true, advance);

    private static void AssertMarkerAdvance(string route, uint defaultThreshold, uint? markerThreshold,
        uint? paragraphMarkThreshold, bool fullOverride, double advance) {
        using WordDocument document = WordDocument.Create();
        document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults =
            new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new RunFonts { Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName },
                new FontSize { Val = "20" }, new Kern { Val = defaultThreshold })));
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        WordListLevel level = list.Numbering.Levels[0];
        level.LevelText = "AV";
        level.LevelJustification = WordListLevelAlignment.Right;
        level.LevelSuffix = WordListLevelSuffix.Nothing;
        var markerProperties = new NumberingSymbolRunProperties(
            new RunFonts { Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName },
            new FontSize { Val = "40" });
        if (markerThreshold.HasValue) markerProperties.Kern = new Kern { Val = markerThreshold.Value };
        level.OpenXmlElement.NumberingSymbolRunProperties = markerProperties;
        if (fullOverride) {
            var replacement = (Level)level.OpenXmlElement.CloneNode(true);
            level.OpenXmlElement.NumberingSymbolRunProperties.Kern = new Kern { Val = markerThreshold == 0 ? 2U : 0U };
            document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart!.Numbering!
                .Elements<NumberingInstance>().Single().Append(new LevelOverride(replacement) { LevelIndex = 0 });
        }
        WordParagraph paragraph = route switch {
            "cell" => document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0],
            "header" => document.HeaderDefaultOrCreate.AddParagraph("B"),
            "footer" => document.FooterDefaultOrCreate.AddParagraph("B"),
            _ => list.AddItem("B")
        };
        paragraph.Text = "B";
        paragraph.FontFamily = "Courier";
        paragraph.FontSize = 10;
        paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.NumberingProperties = new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = list.NumberId });
        if (paragraphMarkThreshold.HasValue)
            paragraph._paragraph.ParagraphProperties.ParagraphMarkRunProperties =
                new ParagraphMarkRunProperties(new Kern { Val = paragraphMarkThreshold.Value });
        if (route is "header" or "footer") document.AddParagraph("B").FontFamily = "Courier";
        var options = new PdfOptions { DefaultFontSize = 10D };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithPairPositioning('A', 'V')));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = options, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "V").ToArray();
        Assert.Equal("AV", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(advance, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 3);
    }
}
