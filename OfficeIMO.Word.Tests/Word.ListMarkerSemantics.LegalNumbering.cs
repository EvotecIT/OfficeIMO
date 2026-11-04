using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(true, "decimalZero", "upperRoman", "Section 1.01", "Section 1.02", "Section 2.01")]
    [InlineData(false, "decimalZero", "upperRoman", "Section I.01", "Section I.02", "Section II.01")]
    [InlineData(true, "upperLetter", "upperRoman", "Section 1.1", "Section 1.2", "Section 2.1")]
    public void LegalListMarkersUseDecimalReferencesAndRetainArabicZeroPadding(
        bool legal, string sectionFormat, string articleFormat, string first, string second, string restarted) {
        using var source = WordDocument.Create();
        var list = source.AddList(WordListStyle.ArticleSections);
        list.AddItem("First article");
        list.AddItem("First section", 1);
        list.AddItem("Second section", 1);
        list.AddItem("Subitem", 2);
        list.AddItem("Second article");
        list.AddItem("Restarted section", 1);
        var definition = source.OpenXmlDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering
            .Elements<AbstractNum>().Single(item => item.AbstractNumberId?.Value == list.Numbering.AbstractNumberId);
        definition.Elements<Level>().Single(level => level.LevelIndex?.Value == 0).NumberingFormat!.Val = new NumberFormatValues(articleFormat);
        var section = definition.Elements<Level>().Single(level => level.LevelIndex?.Value == 1);
        section.NumberingFormat!.Val = new NumberFormatValues(sectionFormat);
        section.IsLegalNumberingStyle = new IsLegalNumberingStyle { Val = legal };

        using var reopened = WordDocument.Load(new System.IO.MemoryStream(source.ToBytes()));
        var markers = WordDocumentTraversal.BuildListMarkers(reopened);
        string Marker(string text) => markers[reopened.Paragraphs.Single(paragraph => paragraph.Text == text)].Marker;
        Assert.Equal(first, Marker("First section"));
        Assert.Equal(second, Marker("Second section"));
        Assert.Equal(restarted, Marker("Restarted section"));
        Assert.Equal(articleFormat == "none" ? "" : "Article I.", Marker("First article"));
        Assert.Equal("(a)", Marker("Subitem"));
        Assert.Empty(reopened.ValidateDocument());
    }
}
