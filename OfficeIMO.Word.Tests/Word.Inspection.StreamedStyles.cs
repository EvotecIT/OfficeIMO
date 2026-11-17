using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void InspectionOfLoadedDocumentReadsStylesWithoutLoadingStylesPartAndMatchesLoadedDom() {
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (WordDocument source = WordDocument.Create(stream)) {
                WordList list = source.AddCustomList();
                list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                list.AddItem("Direct item");
                int numberId = list._numberId;
                Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new SmallCaps());
                styles.Append(
                    new Style(new StyleName { Val = "Streamed base" }, new StyleRunProperties(new Caps())) {
                        Type = StyleValues.Paragraph, StyleId = "StreamedBase"
                    },
                    new Style(new StyleName { Val = "Streamed numbered" }, new BasedOn { Val = "StreamedBase" },
                        new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = numberId })),
                        new StyleRunProperties(new SmallCaps { Val = false })) {
                        Type = StyleValues.Paragraph, StyleId = "StreamedNumbered"
                    },
                    new Style(new StyleName { Val = "Streamed character" }, new StyleRunProperties(new Caps { Val = false })) {
                        Type = StyleValues.Character, StyleId = "StreamedCharacter"
                    });
                source.AddParagraph("Styled item").SetStyleId("StreamedNumbered");
                WordParagraph mixed = source.AddParagraph("Inherited caps").SetStyleId("StreamedNumbered");
                mixed.AddText(" character reset").SetCharacterStyleId("StreamedCharacter");
                source.AddParagraph("Plain");
                source.Save();
            }
            package = stream.ToArray();
        }

        string[] streamed = Describe(loadStylesFirst: false, out bool stylesLoaded);
        Assert.False(stylesLoaded);
        Assert.Equal(Describe(loadStylesFirst: true, out _), streamed);
        // Numbering and caps come through the style chain (base caps, document-default small caps) and the character style.
        Assert.Equal(new[] {
            "ListParagraph|1.|SmallCaps", "Streamed numbered|2.|Caps", "Streamed numbered|3.|Caps",
            "Streamed numbered|3.|", "||SmallCaps"
        }, streamed);

        string[] Describe(bool loadStylesFirst, out bool loaded) {
            using var input = new MemoryStream(package, writable: false);
            using WordDocument document = WordDocument.Load(input, new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            var stylesPart = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!;
            if (loadStylesFirst) _ = stylesPart.Styles;
            string[] result = document.Inspect(includeImageBytes: false).Sections
                .SelectMany(section => section.Elements.OfType<WordParagraphInfo>())
                .SelectMany(paragraph => paragraph.Runs.Select(run =>
                    $"{paragraph.StyleName}|{paragraph.ListMarker}|{run.CapsStyle}"))
                .ToArray();
            loaded = stylesPart.IsRootElementLoaded;
            return result;
        }
    }
}
