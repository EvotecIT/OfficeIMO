using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfLigatureTests {
    public static IEnumerable<object[]> StoryCases() {
        foreach (string story in new[] { "body", "cell", "header", "footer", "link" }) {
            foreach (string? setting in new[] { null, "none", "standard" })
                yield return new object[] { story, setting!, setting == "standard" };
        }
    }

    [Theory]
    [MemberData(nameof(StoryCases))]
    public void DocumentLigatureDefaultsAndExplicitSettingsReachTextStories(string story, string? setting, bool expectedLigature) =>
        AssertLigature(Render(story, null, setting, "15", "1"), expectedLigature);

    [Theory]
    [InlineData("standard", null, true)]
    [InlineData("standard", "none", false)]
    [InlineData("none", "standard", true)]
    [InlineData("none", null, false)]
    public void LigaturesInheritFromDocumentDefaultsAndRespectRunOverrides(string inherited, string? direct, bool expectedLigature) =>
        AssertLigature(Render("body", inherited, direct, "15", "1"), expectedLigature);

    [Theory]
    [InlineData(null, null, false)]
    [InlineData("14", null, false)]
    [InlineData("15", null, true)]
    [InlineData("15", "0", false)]
    [InlineData("14", "1", true)]
    public void CompatibilitySettingsControlSourceLigatures(string? mode, string? enabled, bool expectedLigature) =>
        AssertLigature(Render("body", null, "standard", mode, enabled), expectedLigature);

    [Theory]
    [InlineData("standard", "liga")]
    [InlineData("contextual", "clig")]
    [InlineData("discretional", "dlig")]
    [InlineData("historical", "hlig")]
    [InlineData("all", "dlig")]
    [InlineData("standardDiscretionalHistorical", "hlig")]
    public void LigatureCategoriesSelectTheirSourceOpenTypeFeatures(string setting, string feature) =>
        AssertLigature(Render("body", null, setting, "15", "1", feature), true);

    private static void AssertLigature(byte[] bytes, bool enabled) {
        using var pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("fi", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(enabled ? 1 : 2, letters.Length);
        Assert.Equal(enabled ? 5D : 10D, letters.Sum(letter => letter.Width), 3);
    }

    private static byte[] Render(string story, string? inherited, string? direct, string? mode, string? enabled, string feature = "liga") {
        using WordDocument document = WordDocument.Create();
        var part = document._wordprocessingDocument.MainDocumentPart!;
        var defaults = new W.RunPropertiesBaseStyle(new W.RunFonts {
            Ascii = ManagedTextShapingTestAssets.FamilyName, HighAnsi = ManagedTextShapingTestAssets.FamilyName
        }, new W.FontSize { Val = "20" });
        if (inherited != null) defaults.Append(Ligatures(inherited));
        part.StyleDefinitionsPart!.Styles!.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(defaults));
        W.Settings settings = part.DocumentSettingsPart!.Settings!;
        settings.GetFirstChild<W.Compatibility>()?.Remove();
        var compatibility = new W.Compatibility();
        if (mode != null) compatibility.Append(new W.CompatibilitySetting {
            Name = W.CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = mode
        });
        if (enabled != null) compatibility.Append(new W.CompatibilitySetting {
            Name = new EnumValue<W.CompatSettingNameValues> { InnerText = "enableOpenTypeFeatures" },
            Uri = "http://schemas.microsoft.com/office/word", Val = enabled
        });
        settings.Append(compatibility);
        WordParagraph paragraph;
        if (story == "cell") {
            var table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = null;
            paragraph = table.Rows[0].Cells[0].Paragraphs[0];
            paragraph.Text = "fi";
        } else if (story is "header" or "footer") {
            document.AddParagraph(" ");
            WordHeaderFooter runningStory = story == "header" ? document.HeaderDefaultOrCreate : document.FooterDefaultOrCreate;
            paragraph = runningStory.AddParagraph("fi");
        } else {
            paragraph = document.AddParagraph(story == "link" ? string.Empty : "fi");
            if (story == "link") paragraph = paragraph.AddHyperLink("fi", new Uri("https://example.com/"));
        }
        paragraph.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        paragraph.FontSize = 10;
        if (direct != null) paragraph._paragraph.Descendants<W.Run>().Last(run => run.InnerText == "fi").RunProperties!.Append(Ligatures(direct));
        var options = new PdfOptions { DefaultFontSize = 10D };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i', featureTag: feature, scriptTag: "latn")));
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        return loaded.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = options, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
    }

    private static W14.Ligatures Ligatures(string value) => new() {
        Val = new EnumValue<W14.LigaturesValues> { InnerText = value }
    };
}
