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

    [Fact]
    public void LevelsAddedToAListKeepTheirIndentationAndFont() {
        using var document = WordDocument.Create();
        WordList list = document.AddCustomList();
        WordListLevel level = list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.BulletSquare));

        // The level is built on its own and becomes the numbering part's element when it is added; the part writes
        // w:left, as Word does, and the getters read it back.
        Assert.Equal(720, level.IndentationLeft.ToTwips());
        Assert.Equal(360, level.IndentationHanging.ToTwips());
        Level added = level.OpenXmlLevel();
        Assert.Same(added, list.Numbering.Levels[0].OpenXmlLevel());
        Assert.Equal("720", added.PreviousParagraphProperties!.Indentation!.Left!.Value);
        Assert.Null(added.PreviousParagraphProperties.Indentation.Start);
        Assert.Equal("Wingdings", added.NumberingSymbolRunProperties!.GetFirstChild<RunFonts>()!.Ascii!.Value);
        Assert.Equal("§", level.LevelText);
    }

    [Fact]
    public void FailedLevelEditsLeaveTheListAsItWas() {
        using var document = WordDocument.Create();
        WordList list = document.AddCustomList();
        WordListLevel level = list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));

        // A level a list already holds can't be added again, and trying changes nothing.
        Assert.Throws<InvalidOperationException>(() => list.Numbering.AddLevel(level));
        Assert.Equal(new[] { 0 }, list.Numbering.Levels.Select(existing => existing.LevelIndex));

        // An undefined format or an indentation beyond the range Word stores is rejected before the level changes.
        Assert.ThrowsAny<ArgumentException>(() => level.NumberFormat = (WordNumberFormat)(-1));
        Assert.Equal(WordNumberFormat.None, level.NumberFormat);
        // 200 million points is a valid length, but its four billion twips don't fit the attribute.
        OfficeIMO.OfficeLength tooWide = OfficeIMO.OfficeLength.Points(200_000_000);
        Assert.Throws<OverflowException>(() => level.IndentationLeft = tooWide);
        Assert.Null(level.OpenXmlLevel().PreviousParagraphProperties);
    }

    [Fact]
    public void LevelSettingsInsideExtensionMarkupAreNotTheLevels() {
        using var document = WordDocument.Create();
        WordList list = document.AddCustomList();
        WordListLevel level = list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        Level sdkLevel = level.OpenXmlLevel();
        sdkLevel.GetFirstChild<StartNumberingValue>()!.Remove();
        const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var extension = new OpenXmlUnknownElement("x", "levelExtension", "urn:officeimo:test");
        extension.InnerXml = $"<w:start xmlns:w=\"{WordNamespace}\" w:val=\"9\"/><w:pPr xmlns:w=\"{WordNamespace}\"><w:ind w:left=\"99\"/></w:pPr>";
        sdkLevel.AppendChild(extension);
        string extensionMarkup = extension.OuterXml;

        // The level reads and edits its own w:start and w:pPr/w:ind, which it adds, and leaves the extension alone.
        Assert.Equal(1, level.StartNumberingValue);
        Assert.Equal(0, level.IndentationLeft.ToTwips());
        level.StartNumberingValue = 3;
        level.IndentationLeft = OfficeIMO.OfficeLength.Points(36);
        Assert.Equal("3", sdkLevel.GetFirstChild<StartNumberingValue>()!.Val!.InnerText);
        Assert.Equal("720", sdkLevel.PreviousParagraphProperties!.Indentation!.Left!.Value);
        Assert.Equal(extensionMarkup, extension.OuterXml);
    }
}
