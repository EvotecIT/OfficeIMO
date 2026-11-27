using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void BuiltInListDefinitionsAreTypedDefinitionsUnderTheirOwnTemplateCodes() {
        foreach (WordListStyle style in Enum.GetValues(typeof(WordListStyle))) {
            var definition = new AbstractNum(WordListStyles.GetDefinitionMarkup(style));

            // Every element reads back as the SDK's typed class: a misspelled or misplaced element would be unknown.
            Assert.Empty(definition.Descendants<OpenXmlUnknownElement>());
            Assert.Equal(style, WordListStyles.MatchStyle(definition.GetFirstChild<TemplateCode>()!.Val!.Value));
            Level[] levels = definition.Elements<Level>().ToArray();
            Assert.Equal(style == WordListStyle.Custom ? 0 : 9, levels.Length);
            Assert.Equal(Enumerable.Range(0, levels.Length), levels.Select(level => level.LevelIndex!.Value));
        }

        // Producers that write lowercase hexadecimal still match the built-in style.
        Assert.Equal(WordListStyle.Numbered, WordListStyles.MatchStyle("25be4216"));
        Assert.Equal(WordListStyle.Custom, WordListStyles.MatchStyle(null));

        var picture = new AbstractNum(WordListStyles.PictureBulletMarkup);
        Assert.Empty(picture.Descendants<OpenXmlUnknownElement>());
        Assert.Single(picture.Descendants<LevelPictureBulletId>());
        Assert.Equal(WordListStyle.Custom, WordListStyles.MatchStyle(picture.GetFirstChild<TemplateCode>()!.Val!.Value));
    }

    [Fact]
    public void AddedListDefinitionsGetTheirOwnIdentitiesAndKeepLeftIndentation() {
        using var document = WordDocument.Create();
        foreach (WordListStyle style in new[] { WordListStyle.Headings111, WordListStyle.Numbered, WordListStyle.Numbered, WordListStyle.Custom, WordListStyle.Custom }) {
            document.AddList(style).AddItem(style + " item");
        }

        Numbering numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        AbstractNum[] definitions = numbering.Elements<AbstractNum>().ToArray();
        Assert.Equal(5, definitions.Length);

        // Two lists of the same style are separate definitions to Word, so each gets its own identifier and nsid, and
        // custom definitions their own template code.
        Assert.Equal(definitions.Length, definitions.Select(definition => definition.AbstractNumberId!.Value).Distinct().Count());
        Assert.Equal(definitions.Length, definitions.Select(definition => definition.Nsid!.Val!.Value).Distinct().Count());
        Assert.DoesNotContain(definitions, definition => definition.Nsid!.Val!.Value == "00000000");
        string[] customCodes = definitions.Skip(3).Select(definition => definition.GetFirstChild<TemplateCode>()!.Val!.Value!).ToArray();
        Assert.NotEqual(customCodes[0], customCodes[1]);
        Assert.All(customCodes, code => Assert.False(WordListStyles.IsPlaceholderTemplateCode(code)));

        Indentation[] indentations = numbering.Descendants<Indentation>().ToArray();
        Assert.NotEmpty(indentations);
        Assert.All(indentations, indentation => {
            Assert.NotNull(indentation.Left);
            Assert.Null(indentation.Start);
        });
    }
}
