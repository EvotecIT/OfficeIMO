using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordNumberingHandleTests {
    [Fact]
    public void AddedNumberingSettingsEditTheirSectionAndSurviveReconfiguration() {
        using var document = WordDocument.Create();
        document.AddParagraph("First section");
        var pages = document.AddPageNumbering(3, WordNumberFormat.UpperRoman);
        var footnotes = document.AddFootnoteProperties(startNumber: 2);
        var endnotes = document.AddEndnoteProperties(startNumber: 4);
        pages.StartNumber = 7;
        footnotes.StartNumber = 5;
        endnotes.NumberingFormat = WordNumberFormat.LowerRoman;
        document.AddPageNumbering(11);
        Assert.Equal(11, pages.StartNumber);
        pages.NumberingFormat = WordNumberFormat.UpperLetter;
        var second = document.AddSection();
        var secondPages = second.AddPageNumbering(21);
        secondPages.StartNumber = 23;
        var secondEndnotes = second.AddEndnoteProperties(position: WordEndnotePosition.SectionEnd);
        Assert.Equal(WordEndnotePosition.SectionEnd, endnotes.Position);
        secondEndnotes.Position = null;
        Assert.Equal(WordEndnotePosition.DocumentEnd, endnotes.Position);
        Assert.Equal(4, endnotes.StartNumber);

        byte[] bytes = document.ToBytes();
        using var loaded = WordDocument.Load(new MemoryStream(bytes));
        Assert.Equal(11, loaded.PageNumberSettings.StartNumber);
        Assert.Equal(WordNumberFormat.UpperLetter, loaded.PageNumberSettings.NumberingFormat);
        Assert.Equal(5, loaded.FootnoteSettings.StartNumber);
        Assert.Equal(4, loaded.EndnoteSettings.StartNumber);
        Assert.Equal(WordNumberFormat.LowerRoman, loaded.EndnoteSettings.NumberingFormat);
        Assert.Equal(23, loaded.Sections[1].PageNumberSettings.StartNumber);
        var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365)
            .Validate(loaded._wordprocessingDocument).ToArray();
        Assert.True(errors.Length == 0, string.Join(System.Environment.NewLine, errors.Select(error => error.Description)));
    }
}
