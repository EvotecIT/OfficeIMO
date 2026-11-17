using DocumentFormat.OpenXml;
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

    [Fact]
    public void InspectionHonorsMarkupCompatibilityProcessingOfTheStylesPart() {
        byte[] package;
        using (var stream = new MemoryStream()) {
            using (WordDocument source = WordDocument.Create(stream)) {
                Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                if (styles.LookupNamespace("mc") == null) styles.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                if (styles.LookupNamespace("w14") == null) styles.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                styles.Append(new AlternateContent(
                    new AlternateContentChoice(new Style(new StyleName { Val = "Chosen name" }) {
                        Type = StyleValues.Paragraph, StyleId = "CompatibilityStyle"
                    }) { Requires = "w14" },
                    new AlternateContentFallback(new Style(new StyleName { Val = "Fallback name" }) {
                        Type = StyleValues.Paragraph, StyleId = "CompatibilityStyle"
                    })));
                source.AddParagraph("Alternate").SetStyleId("CompatibilityStyle");
                source.Save();
            }
            package = stream.ToArray();
        }

        // Processing selects the Office 2010 choice while loading; without it the alternate content is opaque.
        Assert.Equal("Chosen name", StyleName(OfficeOpenXmlMarkupCompatibilityMode.ProcessAllParts, loadStylesFirst: false));
        Assert.Equal(StyleName(OfficeOpenXmlMarkupCompatibilityMode.ProcessAllParts, loadStylesFirst: true),
            StyleName(OfficeOpenXmlMarkupCompatibilityMode.ProcessAllParts, loadStylesFirst: false));
        Assert.Equal(StyleName(OfficeOpenXmlMarkupCompatibilityMode.NoProcess, loadStylesFirst: true),
            StyleName(OfficeOpenXmlMarkupCompatibilityMode.NoProcess, loadStylesFirst: false));

        string? StyleName(OfficeOpenXmlMarkupCompatibilityMode mode, bool loadStylesFirst) {
            using var input = new MemoryStream(package, writable: false);
            using WordDocument document = WordDocument.Load(input, new WordLoadOptions {
                AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly,
                OpenSettings = new OfficeOpenXmlLoadSettings {
                    MarkupCompatibilityMode = mode,
                    MarkupCompatibilityTargetVersion = OfficeOpenXmlFileFormatVersion.Office2010
                }
            });
            if (loadStylesFirst) _ = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles;
            return Assert.Single(Assert.Single(document.Inspect(includeImageBytes: false).Sections)
                .Elements.OfType<WordParagraphInfo>()).StyleName;
        }
    }
}
