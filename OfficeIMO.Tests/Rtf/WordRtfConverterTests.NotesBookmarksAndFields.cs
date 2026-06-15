using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Word_To_Rtf_Bridge_Carries_Footnotes() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Body text");
        paragraph.AddFootNote("Footnote text");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        RtfRun run = Assert.Single(rtfParagraph.Runs);
        Assert.Equal("Body text", run.Text);
        Assert.NotNull(run.Note);
        Assert.Equal(RtfNoteKind.Footnote, run.Note!.Kind);
        Assert.Equal("Footnote text", run.Note.ToPlainText());
        Assert.Contains(@"{\footnote", rtf, StringComparison.Ordinal);
        Assert.Contains("Footnote text", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_Endnotes() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Body text");
        paragraph.AddEndNote("Endnote text");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        RtfRun run = Assert.Single(rtfParagraph.Runs);
        Assert.Equal("Body text", run.Text);
        Assert.NotNull(run.Note);
        Assert.Equal(RtfNoteKind.Endnote, run.Note!.Kind);
        Assert.Equal("Endnote text", run.Note.ToPlainText());
        Assert.Contains(@"{\endnote", rtf, StringComparison.Ordinal);
        Assert.Contains("Endnote text", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Footnotes_With_Rich_Text() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        RtfRun bodyRun = paragraph.AddText("Body text");
        var note = new RtfNote(RtfNoteKind.Footnote);
        RtfParagraph noteParagraph = note.AddParagraph("Footnote ");
        noteParagraph.AddText("bold").SetBold();
        bodyRun.SetNote(note);

        using WordDocument word = rtfDocument.ToWordDocument();

        WordFootNote footNote = Assert.Single(word.FootNotes);
        Assert.NotNull(footNote.ParentParagraph);
        Assert.Equal("Body text", footNote.ParentParagraph!.Text);
        List<WordParagraph> noteParagraphs = footNote.Paragraphs!.Where(item => !string.IsNullOrEmpty(item.Text)).ToList();
        Assert.Equal("Footnote bold", string.Concat(noteParagraphs.Select(item => item.Text)));
        Assert.Contains(noteParagraphs, item => item.Text == "bold" && item.Bold);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Endnotes_With_Rich_Text() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        RtfRun bodyRun = paragraph.AddText("Body text");
        var note = new RtfNote(RtfNoteKind.Endnote);
        RtfParagraph noteParagraph = note.AddParagraph("Endnote ");
        noteParagraph.AddText("bold").SetBold();
        bodyRun.SetNote(note);

        using WordDocument word = rtfDocument.ToWordDocument();

        WordEndNote endNote = Assert.Single(word.EndNotes);
        Assert.NotNull(endNote.ParentParagraph);
        Assert.Equal("Body text", endNote.ParentParagraph!.Text);
        List<WordParagraph> noteParagraphs = endNote.Paragraphs!.Where(item => !string.IsNullOrEmpty(item.Text)).ToList();
        Assert.Equal("Endnote bold", string.Concat(noteParagraphs.Select(item => item.Text)));
        Assert.Contains(noteParagraphs, item => item.Text == "bold" && item.Bold);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_Bookmarks_In_Inline_Order() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("Target").AddBookmark("Anchor");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        RtfParagraph paragraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Equal("Target", paragraph.ToPlainText());
        Assert.Collection(paragraph.Inlines,
            inline => Assert.Equal("Target", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfBookmarkMarker marker = Assert.IsType<RtfBookmarkMarker>(inline);
                Assert.Equal(RtfBookmarkMarkerKind.Start, marker.Kind);
                Assert.Equal("Anchor", marker.Name);
            },
            inline => {
                RtfBookmarkMarker marker = Assert.IsType<RtfBookmarkMarker>(inline);
                Assert.Equal(RtfBookmarkMarkerKind.End, marker.Kind);
                Assert.Equal("Anchor", marker.Name);
            });
        Assert.Contains(@"{\*\bkmkstart Anchor}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\bkmkend Anchor}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Bookmarks() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        paragraph.AddBookmarkStart("Anchor");
        paragraph.AddText("Target");
        paragraph.AddBookmarkEnd("Anchor");

        using WordDocument word = rtfDocument.ToWordDocument();

        WordBookmark bookmark = Assert.Single(word.Bookmarks);
        Assert.Equal("Anchor", bookmark.Name);
        Assert.Contains(word.Paragraphs, item => item.Text == "Target");
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_SimpleFields() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        var simpleField = new SimpleField { Instruction = @" PAGE \* MERGEFORMAT " };
        simpleField.Append(new Run(new Text("1")));
        paragraph._paragraph.Append(simpleField);

        RtfDocument rtfDocument = word.ToRtfDocument();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        RtfField field = Assert.IsType<RtfField>(Assert.Single(rtfParagraph.Inlines));
        Assert.Equal(@"PAGE \* MERGEFORMAT", field.Instruction);
        Assert.Equal("1", field.ToPlainText());
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_ComplexFields() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddField(WordFieldType.Page, advanced: true);

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        RtfField field = Assert.IsType<RtfField>(Assert.Single(rtfParagraph.Inlines));
        Assert.Contains("PAGE", field.Instruction, StringComparison.Ordinal);
        Assert.Equal("[Document Page]", field.ToPlainText());
        Assert.Contains(@"{\field{\*\fldinst", rtf, StringComparison.Ordinal);
        Assert.Contains("PAGE", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_GenericFields_As_SimpleFields() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph("Page ");
        RtfField field = paragraph.AddField(@"PAGE \* MERGEFORMAT");
        field.AddText("1");

        using WordDocument word = rtfDocument.ToWordDocument();

        SimpleField simpleField = Assert.Single(word._document.Body!.Descendants<SimpleField>());
        Assert.Equal(@"PAGE \* MERGEFORMAT", simpleField.Instruction?.Value);
        Assert.Equal("1", string.Concat(simpleField.Descendants<Text>().Select(text => text.Text)));
    }

    private static string GetCellText(WordTable table, int rowIndex, int cellIndex) {
        return string.Concat(table.Rows[rowIndex].Cells[cellIndex].Paragraphs.Select(paragraph => paragraph.Text));
    }

    private static byte[] CreateOnePixelPng() {
        return Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
    }
}
