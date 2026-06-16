using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentRichFeatureTests {
    [Fact]
    public void Write_And_Read_Superscript_And_Subscript_Runs() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("2");
        paragraph.AddText("nd").SetSuperscript();
        paragraph.AddText(" H");
        paragraph.AddText("2").SetSubscript();
        paragraph.AddText("O");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\super nd\nosupersub", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sub 2\nosupersub", rtf, StringComparison.Ordinal);
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("2nd H2O", readParagraph.ToPlainText());
        Assert.Contains(readParagraph.Runs, run => run.Text == "nd" && run.VerticalPosition == RtfVerticalPosition.Superscript);
        Assert.Contains(readParagraph.Runs, run => run.Text == "2" && run.VerticalPosition == RtfVerticalPosition.Subscript);
        Assert.Contains(readParagraph.Runs, run => run.Text == "O" && run.VerticalPosition == RtfVerticalPosition.Baseline);
    }

    [Fact]
    public void Write_And_Read_Hidden_Text_Runs() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Visible ");
        paragraph.AddText("Hidden").SetHidden();
        paragraph.AddText(" shown");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\v Hidden\v0", rtf, StringComparison.Ordinal);
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Visible Hidden shown", readParagraph.ToPlainText());
        Assert.Contains(readParagraph.Runs, run => run.Text == "Hidden" && run.Hidden);
        Assert.Contains(readParagraph.Runs, run => run.Text.Contains("shown", StringComparison.Ordinal) && !run.Hidden);
    }

    [Fact]
    public void Write_And_Read_Highlighted_Runs() {
        RtfDocument document = RtfDocument.Create();
        int yellow = document.AddColor(255, 255, 0);
        RtfParagraph paragraph = document.AddParagraph("Normal ");
        paragraph.AddText("highlight").SetHighlightColor(yellow);
        paragraph.AddText(" done");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\colortbl;\red255\green255\blue0;}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\highlight1 highlight", rtf, StringComparison.Ordinal);
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Normal highlight done", readParagraph.ToPlainText());
        Assert.Contains(readParagraph.Runs, run => run.Text == "highlight" && run.HighlightColorIndex == yellow);
    }

    [Fact]
    public void Write_And_Read_Run_Revisions() {
        RtfDocument document = RtfDocument.Create();
        int alice = document.AddRevisionAuthor("Alice");
        int bob = document.AddRevisionAuthor("Bob");
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Base ");
        paragraph.AddText("Inserted").SetInsertedRevision(alice, 123);
        paragraph.AddText(" ");
        paragraph.AddText("Removed").SetDeletedRevision(bob);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\revtbl{Alice;}{Bob;}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\revised \revauth0 \revdttm123 Inserted", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\deleted \revauth1 Removed", rtf, StringComparison.Ordinal);
        Assert.Collection(read.Document.RevisionAuthors,
            author => Assert.Equal("Alice", author.Name),
            author => Assert.Equal("Bob", author.Name));
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Contains(readParagraph.Runs, run => run.Text == "Inserted" && run.RevisionKind == RtfRevisionKind.Inserted && run.RevisionAuthorIndex == alice && run.RevisionTimestampValue == 123);
        Assert.Contains(readParagraph.Runs, run => run.Text == "Removed" && run.RevisionKind == RtfRevisionKind.Deleted && run.RevisionAuthorIndex == bob);
    }

    [Fact]
    public void Write_And_Read_Revision_Save_Id_Table() {
        RtfDocument document = RtfDocument.Create();
        document.SetRevisionRootSaveId(7)
            .AddRevisionSaveId(15)
            .AddRevisionSaveId(1024)
            .AddRevisionSaveId(65535);
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.SetRevisionSaveId(20);
        paragraph.AddText("Base ");
        paragraph.AddText("Revised").SetRevisionSaveIds(character: 30, insertion: 40, deletion: 50);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\rsidtbl\rsidroot7\rsid15\rsid1024\rsid65535}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pard\pararsid20\ql Base \charrsid30 \insrsid40 \delrsid50 Revised\par", rtf, StringComparison.Ordinal);
        Assert.Equal(7, read.Document.RevisionRootSaveId);
        Assert.Equal(new[] { 15, 1024, 65535 }, read.Document.RevisionSaveIds);
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal(20, readParagraph.RevisionSaveId);
        Assert.Equal("Base Revised", readParagraph.ToPlainText());
        RtfRun revised = readParagraph.Runs.Single(run => run.Text == "Revised");
        Assert.Equal(30, revised.CharacterRevisionSaveId);
        Assert.Equal(40, revised.InsertionRevisionSaveId);
        Assert.Equal(50, revised.DeletionRevisionSaveId);
    }

    [Fact]
    public void Write_And_Read_Annotation_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfRun run = document.AddParagraph("Target").Runs[0];
        var note = new RtfNote(RtfNoteKind.Annotation) {
            Id = "c1",
            Author = "Alice",
            Created = new DateTime(2026, 1, 2, 3, 4, 5)
        };
        note.AddParagraph("Review note");
        run.SetNote(note);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\annotation{\*\atnid c1}{\*\atnauthor Alice}{\*\atntime\yr2026\mo1\dy2\hr3\min4\sec5}\chatn", rtf, StringComparison.Ordinal);
        RtfNote? noteFromRead = Assert.Single(read.Document.Paragraphs).Runs[0].Note;
        Assert.NotNull(noteFromRead);
        RtfNote readNote = noteFromRead!;
        Assert.Equal(RtfNoteKind.Annotation, readNote.Kind);
        Assert.Equal("c1", readNote.Id);
        Assert.Equal("Alice", readNote.Author);
        Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), readNote.Created);
        Assert.Equal("Review note", readNote.ToPlainText());
    }

    [Fact]
    public void Write_And_Read_Explicit_Inline_Breaks() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Before");
        paragraph.AddLineBreak();
        paragraph.AddText("Line");
        paragraph.AddSoftLineBreak();
        paragraph.AddText("SoftLine");
        paragraph.AddPageBreak();
        paragraph.AddText("Page");
        paragraph.AddSoftPageBreak();
        paragraph.AddText("SoftPage");
        paragraph.AddColumnBreak();
        paragraph.AddText("Column");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"Before\line Line\softline SoftLine\page Page\softpage SoftPage\column Column", rtf, StringComparison.Ordinal);
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Before" + Environment.NewLine + "Line" + Environment.NewLine + "SoftLine\fPage\fSoftPage\vColumn", readParagraph.ToPlainText());
        Assert.Collection(readParagraph.Inlines,
            inline => Assert.Equal("Before", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Line, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("Line", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.SoftLine, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("SoftLine", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Page, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("Page", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.SoftPage, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("SoftPage", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBreakKind.Column, Assert.IsType<RtfBreak>(inline).Kind),
            inline => Assert.Equal("Column", Assert.IsType<RtfRun>(inline).Text));
    }

    [Fact]
    public void Write_And_Read_Generated_Text_Controls() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Page ");
        paragraph.AddPageNumber();
        paragraph.AddText(" Section ");
        paragraph.AddSectionNumber();
        paragraph.AddText(" Date ");
        paragraph.AddCurrentDate();
        paragraph.AddText(" Long ");
        paragraph.AddCurrentDateLong();
        paragraph.AddText(" Short ");
        paragraph.AddCurrentDateAbbreviated();
        paragraph.AddText(" Time ");
        paragraph.AddCurrentTime();

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\chpgn", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sectnum", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chdate", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chdpl", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chdpa", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chtime", rtf, StringComparison.Ordinal);
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Collection(readParagraph.Inlines,
            inline => Assert.Equal("Page ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.PageNumber, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal(" Section ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.SectionNumber, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal(" Date ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentDate, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal(" Long ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentDateLong, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal(" Short ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentDateAbbreviated, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal(" Time ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentTime, Assert.IsType<RtfGeneratedText>(inline).Kind));
    }

    [Fact]
    public void Write_And_Read_Tab_Stops_And_Tab_Text() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddTabStop(1440);
        paragraph.AddTabStop(2880, RtfTabAlignment.Right, RtfTabLeader.Dots);
        paragraph.AddTabStop(4320, RtfTabAlignment.Decimal, RtfTabLeader.MiddleDots);
        paragraph.AddText("Name\tAmount\t12.34");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\tx1440", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\tldot\tqr\tx2880", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\tlmdot\tqdec\tx4320", rtf, StringComparison.Ordinal);
        Assert.Contains(@"Name\tab Amount\tab 12.34", rtf, StringComparison.Ordinal);
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Name\tAmount\t12.34", readParagraph.ToPlainText());
        Assert.Collection(readParagraph.TabStops,
            tabStop => {
                Assert.Equal(1440, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Left, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.None, tabStop.Leader);
            },
            tabStop => {
                Assert.Equal(2880, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Right, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.Dots, tabStop.Leader);
            },
            tabStop => {
                Assert.Equal(4320, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Decimal, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.MiddleDots, tabStop.Leader);
            });
    }
}
