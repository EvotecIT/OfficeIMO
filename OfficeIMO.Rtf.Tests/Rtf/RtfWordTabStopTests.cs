using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfWordTabStopTests {
    [Fact]
    public void Rtf_Style_Tabs_Reach_Paragraphs_Cells_And_Notes_Without_Duplicating_Direct_Stops() {
        RtfDocument source = RtfDocument.Create();
        source.AddStyle(7, "Tab parent").AddTabStop(4320, RtfTabAlignment.Right, RtfTabLeader.Dots);
        source.AddStyle(8, "Inherited tabs").BasedOnStyleId = 7;
        source.AddParagraph("Styled\tBody").SetStyle(8);
        source.AddParagraph("Direct\tBody").SetStyle(8).AddTabStop(2880, RtfTabAlignment.Decimal, RtfTabLeader.Hyphen);
        source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Styled\tCell").SetStyle(8);
        var note = new RtfNote(RtfNoteKind.Footnote);
        note.AddParagraph("Styled\tNote").SetStyle(8);
        source.AddParagraph("Anchor").AddText("F").SetNote(note);

        using WordDocument word = source.ToWordDocument();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(word.ToBytes()));
        var main = reopened._wordprocessingDocument.MainDocumentPart!;
        Document mainDocument = Assert.IsType<Document>(main.Document);
        Style parent = Assert.Single(main.StyleDefinitionsPart!.Styles!.Elements<Style>(), item => item.StyleId?.Value == "RtfP7");
        AssertWordTab(Assert.Single(parent.StyleParagraphProperties!.Tabs!.Elements<TabStop>()), 4320, TabStopValues.Right, TabStopLeaderCharValues.Dot);
        Paragraph body = Assert.Single(mainDocument.Descendants<Paragraph>(), item => item.InnerText == "StyledBody");
        Paragraph direct = Assert.Single(mainDocument.Descendants<Paragraph>(), item => item.InnerText == "DirectBody");
        Paragraph cell = Assert.Single(mainDocument.Descendants<Paragraph>(), item => item.InnerText == "StyledCell");
        Paragraph noteBody = Assert.Single(main.FootnotesPart!.Footnotes!.Descendants<Paragraph>(), item => item.InnerText == "StyledNote");
        AssertWordTab(Assert.Single(body.ParagraphProperties!.Tabs!.Elements<TabStop>()), 4320, TabStopValues.Right, TabStopLeaderCharValues.Dot);
        AssertWordTab(Assert.Single(direct.ParagraphProperties!.Tabs!.Elements<TabStop>()), 2880, TabStopValues.Decimal, TabStopLeaderCharValues.Hyphen);
        AssertWordTab(Assert.Single(cell.ParagraphProperties!.Tabs!.Elements<TabStop>()), 4320, TabStopValues.Right, TabStopLeaderCharValues.Dot);
        AssertWordTab(Assert.Single(noteBody.ParagraphProperties!.Tabs!.Elements<TabStop>()), 4320, TabStopValues.Right, TabStopLeaderCharValues.Dot);

        RtfDocument returned = RtfDocument.Read(reopened.ToRtf(new RtfWriteOptions { IncludeGenerator = false })).Document;
        RtfStyle returnedParent = Assert.Single(returned.Styles, item => item.Name == "RtfP7");
        AssertRtfTab(Assert.Single(returnedParent.TabStops), 4320, RtfTabAlignment.Right, RtfTabLeader.Dots);
        RtfParagraph returnedBody = Assert.Single(returned.Paragraphs, item => item.ToPlainText() == "Styled\tBody");
        AssertRtfTab(Assert.Single(returnedBody.TabStops), 4320, RtfTabAlignment.Right, RtfTabLeader.Dots);
        RtfTable returnedTable = Assert.Single(returned.Blocks.OfType<RtfTable>());
        AssertRtfTab(Assert.Single(returnedTable.Rows[0].Cells[0].Paragraphs.Single().TabStops), 4320, RtfTabAlignment.Right, RtfTabLeader.Dots);
        RtfNote returnedNote = Assert.Single(returned.Paragraphs.SelectMany(item => item.Runs), item => item.Note != null).Note!;
        AssertRtfTab(Assert.Single(returnedNote.Paragraphs.Single().TabStops), 4320, RtfTabAlignment.Right, RtfTabLeader.Dots);
    }

    [Fact]
    public void Word_Active_Style_And_Direct_Tabs_RoundTrip_While_Unsupported_Clears_Are_Not_Active_Left_Stops() {
        using WordDocument word = WordDocument.Create();
        var style = new Style { Type = StyleValues.Paragraph, StyleId = "Tabs" };
        style.StyleParagraphProperties = new StyleParagraphProperties(new Tabs(
            new TabStop { Position = 4320, Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot },
            new TabStop { Position = 1440, Val = TabStopValues.Clear }));
        word._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
        WordParagraph paragraph = word.AddParagraph("Direct\tTab");
        paragraph.AddTabStop(2880, WordTabAlignment.Decimal, WordTabLeader.Hyphen);
        paragraph.AddTabStop(720, WordTabAlignment.Clear);
        string before = word._document.OuterXml;
        string styleBefore = style.OuterXml;

        RtfDocument returned = word.ToRtfDocument();

        AssertRtfTab(Assert.Single(Assert.Single(returned.Styles, item => item.Name == "Tabs").TabStops),
            4320, RtfTabAlignment.Right, RtfTabLeader.Dots);
        AssertRtfTab(Assert.Single(Assert.Single(returned.Paragraphs).TabStops), 2880, RtfTabAlignment.Decimal, RtfTabLeader.Hyphen);
        Assert.Equal(before, word._document.OuterXml);
        Assert.Equal(styleBefore, style.OuterXml);
        using WordDocument roundTrip = returned.ToWordDocument();
        Paragraph roundTripParagraph = Assert.Single(roundTrip._document.Body!.Elements<Paragraph>());
        AssertWordTab(Assert.Single(roundTripParagraph.ParagraphProperties!.Tabs!.Elements<TabStop>()),
            2880, TabStopValues.Decimal, TabStopLeaderCharValues.Hyphen);
    }

    private static void AssertWordTab(TabStop stop, int position, TabStopValues alignment, TabStopLeaderCharValues leader) {
        Assert.Equal(position, stop.Position?.Value);
        Assert.Equal(alignment, stop.Val?.Value);
        Assert.Equal(leader, stop.Leader?.Value);
    }

    private static void AssertRtfTab(RtfTabStop stop, int position, RtfTabAlignment alignment, RtfTabLeader leader) {
        Assert.Equal(position, stop.PositionTwips);
        Assert.Equal(alignment, stop.Alignment);
        Assert.Equal(leader, stop.Leader);
    }
}
