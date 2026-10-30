using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.Rtf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfWordTabStopTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReplacedAncestorStyleStopsAreClearedWithoutClearingActivePositions(bool samePosition) {
        RtfDocument source = RtfDocument.Create();
        source.AddStyle(7, "Ancestor tabs").AddTabStop(OfficeLength.Emus(4320L * 635), RtfTabAlignment.Right, RtfTabLeader.Dots);
        RtfStyle child = source.AddStyle(8, "Replacement tabs");
        child.BasedOnStyleId = 7;
        int activePosition = samePosition ? 4320 : 2880;
        child.AddTabStop(activePosition);
        RtfParagraph paragraph = source.AddParagraph("A\tB\tC").SetStyle(8);
        AssertRtfTab(Assert.Single(source.ResolveParagraphFormatting(paragraph).TabStops),
            activePosition, RtfTabAlignment.Left, RtfTabLeader.None);

        using WordDocument word = source.ToWordDocument();
        Paragraph projected = Assert.Single(word.OpenXmlDocument.MainDocumentPart!.Document!.Body!.Elements<Paragraph>());
        TabStop[] stops = projected.ParagraphProperties!.Tabs!.Elements<TabStop>().ToArray();

        AssertWordTab(Assert.Single(stops, stop => stop.Val?.Value != TabStopValues.Clear),
            activePosition, TabStopValues.Left, TabStopLeaderCharValues.None);
        if (samePosition) {
            Assert.DoesNotContain(stops, stop => stop.Val?.Value == TabStopValues.Clear);
        } else {
            Assert.Equal(4320, Assert.Single(stops, stop => stop.Val?.Value == TabStopValues.Clear).Position?.Value);
        }
    }

    [Fact]
    public void DirectTabListKeepsLaterDefaultTabGeometryThroughWordConversion() {
        RtfDocument source = CreateDirectTabSource(out RtfParagraph paragraph, "body");
        paragraph.AddText("A\tB\tC").SetFontSize(12);
        using PdfPigDocument direct = PdfPigDocument.Open((source.ToPdf(new RtfToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        })).ToBytes());
        using WordDocument word = source.ToWordDocument();
        using PdfPigDocument composed = PdfPigDocument.Open((word.ToPdf(new WordPdfExportOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
        })).ToBytes());
        var directLetters = direct.GetPage(1).Letters;
        var composedLetters = composed.GetPage(1).Letters;
        double DirectX(string text) => Assert.Single(directLetters, item => item.Value == text).StartBaseLine.X;
        double ComposedX(string text) => Assert.Single(composedLetters, item => item.Value == text).StartBaseLine.X;

        // The direct list replaces the inherited 216-point right/dot stop.
        // After its 144-point stop, the next ordinary default is 180 points.
        Assert.Equal(180D, DirectX("C") - DirectX("A"), 3);
        Assert.Equal(DirectX("B") - DirectX("A"), ComposedX("B") - ComposedX("A"), 3);
        Assert.Equal(DirectX("C") - DirectX("A"), ComposedX("C") - ComposedX("A"), 3);
        Assert.DoesNotContain(composedLetters, item => item.Value == ".");
    }

    [Theory]
    [InlineData("body")]
    [InlineData("table")]
    [InlineData("footnote")]
    [InlineData("endnote")]
    [InlineData("header")]
    [InlineData("footer")]
    public void DirectTabListClearsReplacedStyleStopsAcrossWordStories(string story) {
        RtfDocument source = CreateDirectTabSource(out RtfParagraph paragraph, story);
        paragraph.AddText("A\tB\tC");
        string original = source.ToRtf();
        AssertRtfTab(Assert.Single(source.ResolveParagraphFormatting(paragraph).TabStops),
            2880, RtfTabAlignment.Left, RtfTabLeader.None);

        using WordDocument word = source.ToWordDocument();
        var main = word.OpenXmlDocument.MainDocumentPart!;
        var roots = new List<DocumentFormat.OpenXml.OpenXmlElement> { main.Document! };
        if (main.FootnotesPart?.Footnotes is { } footnotes) roots.Add(footnotes);
        if (main.EndnotesPart?.Endnotes is { } endnotes) roots.Add(endnotes);
        roots.AddRange(main.HeaderParts.Select(part => part.Header!));
        roots.AddRange(main.FooterParts.Select(part => part.Footer!));
        Paragraph projected = Assert.Single(roots.SelectMany(root => root.Descendants<Paragraph>()),
            item => item.InnerText == "ABC");
        TabStop[] stops = projected.ParagraphProperties!.Tabs!.Elements<TabStop>().ToArray();

        AssertWordTab(Assert.Single(stops, stop => stop.Val?.Value != TabStopValues.Clear),
            2880, TabStopValues.Left, TabStopLeaderCharValues.None);
        Assert.Equal(4320, Assert.Single(stops, stop => stop.Val?.Value == TabStopValues.Clear).Position?.Value);
        Assert.Equal(2, stops.Length);
        Assert.Equal(original, source.ToRtf());
    }

    [Fact]
    public void Rtf_Style_Tabs_Reach_Paragraphs_Cells_And_Notes_Without_Duplicating_Direct_Stops() {
        RtfDocument source = RtfDocument.Create();
        source.AddStyle(7, "Tab parent").AddTabStop(OfficeLength.Emus(4320L * 635), RtfTabAlignment.Right, RtfTabLeader.Dots);
        source.AddStyle(8, "Inherited tabs").BasedOnStyleId = 7;
        source.AddParagraph("Styled\tBody").SetStyle(8);
        source.AddParagraph("Direct\tBody").SetStyle(8).AddTabStop(OfficeLength.Emus(2880L * 635), RtfTabAlignment.Decimal, RtfTabLeader.Hyphen);
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
        AssertWordTab(Assert.Single(direct.ParagraphProperties!.Tabs!.Elements<TabStop>(),
            stop => stop.Val?.Value != TabStopValues.Clear), 2880, TabStopValues.Decimal, TabStopLeaderCharValues.Hyphen);
        Assert.Equal(4320, Assert.Single(direct.ParagraphProperties.Tabs.Elements<TabStop>(),
            stop => stop.Val?.Value == TabStopValues.Clear).Position?.Value);
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
        paragraph.AddTabStop(OfficeLength.Emus(2880L * 635), WordTabAlignment.Decimal, WordTabLeader.Hyphen);
        paragraph.AddTabStop(OfficeLength.Emus(720L * 635), WordTabAlignment.Clear);
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

    private static RtfDocument CreateDirectTabSource(out RtfParagraph paragraph, string story) {
        RtfDocument source = RtfDocument.Create();
        source.Settings.SetDefaultFont(source.AddFont("Arial"));
        RtfStyle parent = source.AddStyle(7, "Tab parent");
        parent.FontSize = 12;
        parent.AddTabStop(OfficeLength.Emus(4320L * 635), RtfTabAlignment.Right, RtfTabLeader.Dots);
        source.AddStyle(8, "Inherited tabs").BasedOnStyleId = parent.Id;
        switch (story) {
            case "table": paragraph = source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph(); break;
            case "header": paragraph = source.AddHeader().AddParagraph(); break;
            case "footer": paragraph = source.AddFooter().AddParagraph(); break;
            case "footnote":
            case "endnote":
                RtfNote note = source.AddNote(story == "endnote" ? RtfNoteKind.Endnote : RtfNoteKind.Footnote);
                source.AddParagraph("Anchor").AddNoteReference(note);
                paragraph = note.AddParagraph();
                break;
            default: paragraph = source.AddParagraph(); break;
        }
        paragraph.SetStyle(8).AddTabStop(OfficeLength.Emus(2880L * 635));
        return source;
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
