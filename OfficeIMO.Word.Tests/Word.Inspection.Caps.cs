using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void InspectionCapsResolveIndependentOverridesWithoutChangingSource() {
        using WordDocument document = WordDocument.Create();
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new Caps(), new SmallCaps());
        styles.Append(
            new Style(new StyleRunProperties(new Caps { Val = false })) {
                Type = StyleValues.Paragraph, StyleId = "InspectionCapsParagraph"
            },
            new Style(new StyleRunProperties(new Caps())) {
                Type = StyleValues.Character, StyleId = "InspectionCapsCharacterBase"
            },
            new Style(new BasedOn { Val = "InspectionCapsCharacterBase" },
                new StyleRunProperties(new SmallCaps { Val = false })) {
                Type = StyleValues.Character, StyleId = "InspectionCapsCharacter"
            });
        WordParagraph paragraph = document.AddParagraph("Inherited").SetStyleId("InspectionCapsParagraph");
        paragraph.AddText("Reset")._run!.RunProperties = new RunProperties(new SmallCaps { Val = false });
        paragraph.AddText("Character").SetCharacterStyleId("InspectionCapsCharacter");
        WordParagraph direct = paragraph.AddText("Direct").SetCharacterStyleId("InspectionCapsCharacter");
        direct._run!.RunProperties!.Caps = new Caps { Val = false };
        paragraph.AddText("Implicit")._run!.RunProperties = new RunProperties(new Caps(), new SmallCaps { Val = false });
        WordHyperLink link = paragraph.AddHyperLink("Linked", new Uri("https://example.com/caps"));
        link.Runs.Single()._run!.RunProperties = new RunProperties(new Caps { Val = false }, new SmallCaps());
        document.AddParagraph("Default");
        string before = document._wordprocessingDocument.MainDocumentPart.Document.OuterXml + styles.OuterXml;

        WordParagraphInfo[] snapshots = Assert.Single(document.Inspect(includeImageBytes: false).Sections)
            .Elements.OfType<WordParagraphInfo>().ToArray();

        Assert.Equal(2, snapshots.Length);
        Assert.Equal(new[] { "Inherited", "Reset", "Character", "Direct", "Implicit", "Linked" },
            snapshots[0].Runs.Select(run => run.Text));
        Assert.Equal(new string?[] { nameof(WordCapsStyle.SmallCaps), null, nameof(WordCapsStyle.Caps), null,
            nameof(WordCapsStyle.Caps), nameof(WordCapsStyle.SmallCaps) }, snapshots[0].Runs.Select(run => run.CapsStyle));
        Assert.True(snapshots[0].Runs[5].IsHyperlink);
        Assert.Equal("https://example.com/caps", snapshots[0].Runs[5].HyperlinkUri);
        Assert.Equal(nameof(WordCapsStyle.Caps), Assert.Single(snapshots[1].Runs).CapsStyle);
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml + styles.OuterXml);
    }
}
