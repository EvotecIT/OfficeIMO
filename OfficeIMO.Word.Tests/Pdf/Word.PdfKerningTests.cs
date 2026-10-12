using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfKerningTests {
    [Theory]
    [InlineData(false, 100, 0)]
    [InlineData(true, 100, 0)]
    [InlineData(false, 150, 20)]
    [InlineData(true, 150, 20)]
    public void DecimalTabsAlignTheSeparatorAfterKerning(bool inTable, int characterScale, int spacingTwips) {
        using WordDocument document = WordDocument.Create();
        document.Settings.DecimalSymbol = ".";
        WordTableCell? cell = inTable ? document.AddTable(1, 1).Rows[0].Cells[0] : null;
        foreach (string text in new[] { "\t1.", "\t." }) {
            WordParagraph paragraph = cell == null ? document.AddParagraph(text) : cell.AddParagraph(text);
            paragraph.FontFamily = ManagedTextShapingTestAssets.FamilyName;
            paragraph.FontSize = 10;
            paragraph.KerningMinimumFontSizePoints = 1D;
            paragraph.CharacterScale = characterScale;
            paragraph._paragraph.Descendants<Run>().Last().RunProperties!.Spacing = new Spacing { Val = spacingTwips };
            paragraph.AddTabStop(2000, WordTabAlignment.Decimal);
        }

        byte[] bytes = document.ToPdfBytes(ConversionOptions(
            FontOptions(ManagedTextShapingTestAssets.CreateFontWithPairPositioning('1', '.', includeSpace: true))));
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is { Length: > 0 } output) {
            Directory.CreateDirectory(output);
            File.WriteAllBytes(Path.Combine(output, $"decimal-kerning-{inTable}-{characterScale}-{spacingTwips}.pdf"), bytes);
        }
        using var pdf = PdfPigDocument.Open(bytes);
        Letter[] separators = pdf.GetPage(1).Letters.Where(letter => letter.Value == ".").ToArray();
        Assert.Equal(2, separators.Length);
        Assert.Equal(separators[0].StartBaseLine.X, separators[1].StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("link")]
    public void KerningThresholdAndExplicitDisableReachEveryTextStory(string route) {
        AssertAdvance(Render(route, "direct", 1D), 4.6D);
        AssertAdvance(Render(route, "direct", 10D), 4.6D);
        AssertAdvance(Render(route, "direct", 10.5D), 5D);
        AssertAdvance(Render(route, "direct", 0D), 5D);
        AssertAdvance(Render(route, "none", null), 5D);
    }

    [Theory]
    [InlineData("document", 4.6D)]
    [InlineData("paragraph", 5D)]
    [InlineData("character", 4.6D)]
    [InlineData("direct", 5D)]
    [InlineData("conditional", 4.6D)]
    public void KerningInheritsThroughStylesAndRespectsExplicitOverrides(string level, double advance) {
        AssertAdvance(Render(level == "conditional" ? "cell" : "body", level, 0D), advance);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("header")]
    public void AdjacentRunsRetainDifferentKerningSettings(string route) {
        Letter[] letters = Render(route, "direct", 1D, mixed: true);
        Assert.Equal("AVAV", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(4.6D, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 3);
        Assert.Equal(5D, letters[3].StartBaseLine.X - letters[2].StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitKerningAndSourceLigaturesReachManagedShaping(bool managedProvider) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("fi");
        paragraph.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        paragraph.FontSize = 10;
        paragraph.KerningMinimumFontSizePoints = 1D;
        paragraph._paragraph.Descendants<Run>().Single().RunProperties!.Append(
            new DocumentFormat.OpenXml.Office2010.Word.Ligatures { Val = DocumentFormat.OpenXml.Office2010.Word.LigaturesValues.Standard });
        document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!
            .GetFirstChild<Compatibility>()!.Append(new CompatibilitySetting {
                Name = new DocumentFormat.OpenXml.EnumValue<CompatSettingNameValues> { InnerText = "enableOpenTypeFeatures" },
                Uri = "http://schemas.microsoft.com/office/word", Val = "1"
            });
        PdfOptions options = FontOptions(ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', scriptTag: "latn"));
        if (managedProvider) options.UseManagedTextShaping();
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(ConversionOptions(options)));
        Letter ligature = Assert.Single(pdf.GetPage(1).Letters);
        Assert.Equal("fi", ligature.Value);
        Assert.Equal(5D, ligature.Width, 3);
    }

    private static void AssertAdvance(Letter[] letters, double expected) {
        Assert.Equal("AV", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(letters[0].StartBaseLine.Y, letters[1].StartBaseLine.Y, 3);
        Assert.Equal(expected, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 3);
    }

    private static Letter[] Render(string route, string level, double? threshold, bool mixed = false) {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        var defaults = new RunPropertiesBaseStyle(new RunFonts {
            Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName
        }, new FontSize { Val = "20" });
        styles.DocDefaults = new DocDefaults(new RunPropertiesDefault(defaults));
        WordTable? table = null;
        WordParagraph paragraph;
        if (route == "cell") {
            table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = null;
            paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "AV";
        } else if (route is "header" or "footer") {
            document.AddParagraph("Body");
            WordHeaderFooter story = route == "header" ? document.HeaderDefaultOrCreate : document.FooterDefaultOrCreate;
            paragraph = story.AddParagraph("AV");
        } else {
            paragraph = document.AddParagraph(route == "link" ? string.Empty : "AV");
            if (route == "link") paragraph = paragraph.AddHyperLink("AV", new Uri("https://example.com/"));
        }
        paragraph.FontSize = 10;
        paragraph.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        // This hierarchy independently tests a threshold inherited from document
        // defaults, a based-on paragraph style, a character style, and a direct reset.
        if (level is "document" or "paragraph" or "character" or "direct") defaults.Kern = new Kern { Val = 2U };
        if (level is "paragraph" or "character") {
            styles.Append(new Style(new StyleRunProperties(new Kern { Val = 24U })) {
                Type = StyleValues.Paragraph, StyleId = "LargeOnly"
            }, new Style(new BasedOn { Val = "LargeOnly" }) { Type = StyleValues.Paragraph, StyleId = "DerivedLargeOnly" });
            paragraph.SetStyleId("DerivedLargeOnly");
        }
        if (level == "character") {
            styles.Append(new Style(new StyleRunProperties(new Kern { Val = 2U })) { Type = StyleValues.Character, StyleId = "Kerned" });
            paragraph._paragraph.Descendants<Run>().Last().RunProperties!.RunStyle = new RunStyle { Val = "Kerned" };
        }
        if (level == "conditional") {
            styles.Append(new Style(new TableStyleProperties(new RunPropertiesBaseStyle(new Kern { Val = 2U })) {
                Type = TableStyleOverrideValues.FirstRow
            }) { Type = StyleValues.Table, StyleId = "KernedHeader" });
            table!._tableProperties!.TableStyle = new TableStyle { Val = "KernedHeader" };
            table.ConditionalFormattingFirstRow = true;
        }
        if (level == "direct") paragraph.KerningMinimumFontSizePoints = threshold;
        if (mixed) {
            WordParagraph following = paragraph.AddText("AV");
            following.FontSize = 10;
            following.FontFamily = ManagedTextShapingTestAssets.FamilyName;
            following.KerningMinimumFontSizePoints = 0D;
        }
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(ConversionOptions(
            FontOptions(ManagedTextShapingTestAssets.CreateFontWithPairPositioning('A', 'V')))));
        return pdf.GetPage(1).Letters.Where(letter => letter.Value is "A" or "V").ToArray();
    }

    private static PdfOptions FontOptions(byte[] font) {
        var options = new PdfOptions { DefaultFontSize = 10D };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName, font));
        return options;
    }

    private static WordToPdfOptions ConversionOptions(PdfOptions options) => new() {
        IncludePageNumbers = false, PdfOptions = options,
        ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
    };
}
