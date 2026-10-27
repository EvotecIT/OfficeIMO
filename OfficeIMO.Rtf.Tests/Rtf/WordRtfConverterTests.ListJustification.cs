using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Theory]
    [InlineData("start", RtfListLevelAlignment.Left)]
    [InlineData("end", RtfListLevelAlignment.Right)]
    [InlineData("center", RtfListLevelAlignment.Center)]
    [InlineData("both", RtfListLevelAlignment.Left)]
    public void WordToRtfAcceptsWordListJustificationTokens(string token, RtfListLevelAlignment expected) {
        using WordDocument word = WordDocument.Create();
        WordList list = word.AddList(WordListStyle.Bulleted);
        list.AddItem("An aligned list item");
        foreach (LevelJustification justification in word._wordprocessingDocument.MainDocumentPart!
            .NumberingDefinitionsPart!.Numbering!.Descendants<LevelJustification>()) {
            justification.SetAttribute(new OpenXmlAttribute("w", "val",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main", token));
        }

        RtfDocument converted = RtfDocument.Parse(word.ToRtf());

        Assert.NotEmpty(converted.ListDefinitions);
        Assert.All(converted.ListDefinitions.SelectMany(definition => definition.Levels),
            level => Assert.Equal(expected, level.Alignment));
        Assert.Contains(converted.Paragraphs, paragraph => paragraph.ToPlainText().Contains("An aligned list item"));
    }
}
