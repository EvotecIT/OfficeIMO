using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void BuiltInListDefinitionsAreTypedDefinitionsWithFreshIdentities() {
        foreach (WordListStyle style in Enum.GetValues(typeof(WordListStyle))) {
            var definition = new AbstractNum(WordListStyles.GetDefinitionMarkup(style, 42));
            var again = new AbstractNum(WordListStyles.GetDefinitionMarkup(style, 42));

            // Every element reads back as the SDK's typed class: a misspelled or misplaced element would be unknown.
            Assert.Empty(definition.Descendants<OpenXmlUnknownElement>());
            Assert.Equal(42, definition.AbstractNumberId!.Value);
            Assert.Equal(style, WordListStyles.MatchStyle(definition.GetFirstChild<TemplateCode>()!.Val!.Value));
            Assert.Equal(style == WordListStyle.Custom ? 0 : 9, definition.Elements<Level>().Count());
            Assert.Equal(Enumerable.Range(0, definition.Elements<Level>().Count()), definition.Elements<Level>().Select(level => level.LevelIndex!.Value));

            // Two lists of the same style are separate definitions to Word, so each gets its own nsid.
            Assert.NotEqual(definition.Nsid!.Val!.Value, again.Nsid!.Val!.Value);
            if (style == WordListStyle.Custom) {
                Assert.NotEqual(definition.GetFirstChild<TemplateCode>()!.Val!.Value, again.GetFirstChild<TemplateCode>()!.Val!.Value);
            }
        }

        var picture = new AbstractNum(WordListStyles.GetPictureBulletMarkup(7, 3));
        Assert.Empty(picture.Descendants<OpenXmlUnknownElement>());
        Assert.Equal(3, picture.AbstractNumberId!.Value);
        Assert.Equal(7, picture.Descendants<LevelPictureBulletId>().Single().Val!.Value);
        Assert.Equal(WordListStyle.Custom, WordListStyles.MatchStyle(picture.GetFirstChild<TemplateCode>()!.Val!.Value));
    }

    [Fact]
    public void AddedListDefinitionsKeepLeftIndentation() {
        using var document = WordDocument.Create();
        document.AddList(WordListStyle.Headings111).AddItem("Heading item");
        document.AddList(WordListStyle.Bulleted).AddItem("Bullet item");

        Numbering numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        Indentation[] indentations = numbering.Descendants<Indentation>().ToArray();
        Assert.NotEmpty(indentations);
        Assert.All(indentations, indentation => {
            Assert.NotNull(indentation.Left);
            Assert.Null(indentation.Start);
        });
    }
}
