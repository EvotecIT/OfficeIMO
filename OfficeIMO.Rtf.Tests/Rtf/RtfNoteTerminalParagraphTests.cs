using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfNoteTerminalParagraphTests {
    [Fact]
    public void ParsedFootnoteRetainsExplicitTerminalParagraphAndItsExactLineState() {
        RtfDocument document = RtfDocument.Read(NativeSource).Document;
        RtfNote note = Assert.Single(document.Notes);
        Assert.Equal(3, note.Paragraphs.Count);
        Assert.Contains("NOTE-MIN-A", note.Paragraphs[0].ToPlainText());
        Assert.Contains("NOTE-EXACT-A", note.Paragraphs[1].ToPlainText());
        AssertTerminalState(note.Paragraphs[2]);

        using WordDocument word = document.ToWordDocument();
        using var stream = new MemoryStream();
        word.Save(stream);
        stream.Position = 0;
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        Footnote persisted = Assert.Single(package.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<Footnote>(), item => item.Id!.Value > 0);
        Paragraph[] paragraphs = persisted.Elements<Paragraph>().ToArray();
        Assert.Equal(3, paragraphs.Length);
        Assert.Empty(paragraphs[2].Descendants<Text>());
        SpacingBetweenLines spacing = paragraphs[2].ParagraphProperties!
            .GetFirstChild<SpacingBetweenLines>()!;
        Assert.Equal("360", spacing.Line!.Value);
        Assert.Equal(LineSpacingRuleValues.Exact, spacing.LineRule!.Value);
        Assert.Equal("0", spacing.Before!.Value);
        Assert.Equal("0", spacing.After!.Value);
    }

    [Fact]
    public void NoteWithoutTerminalParagraphBreakDoesNotGainABlankParagraph() {
        string source = NativeSource.Replace(@"NOTE-EXACT-B\par}", @"NOTE-EXACT-B}");
        RtfNote note = Assert.Single(RtfDocument.Read(source).Document.Notes);
        Assert.Equal(2, note.Paragraphs.Count);
        Assert.Contains("NOTE-EXACT-B", note.Paragraphs[1].ToPlainText());
    }

    [Fact]
    public void NormalizationRetainsTerminalNoteParagraphWithoutGrowingTheStory() {
        RtfDocument document = RtfDocument.Read(NativeSource).Document;
        for (int pass = 0; pass < 3; pass++) {
            RtfNote note = Assert.Single(document.Notes);
            Assert.Equal(3, note.Paragraphs.Count);
            AssertTerminalState(note.Paragraphs[2]);
            document = RtfDocument.Read(document.ToRtf()).Document;
        }
    }

    [Theory]
    [InlineData(RtfNoteKind.Footnote)]
    [InlineData(RtfNoteKind.Endnote)]
    public void AuthoredEmptyLastNoteParagraphDoesNotGrowDuringNormalization(RtfNoteKind kind) {
        RtfDocument document = RtfDocument.Create();
        var note = new RtfNote(kind);
        note.AddParagraph("Content");
        RtfParagraph last = note.AddParagraph();
        last.LineSpacingTwips = -360;
        last.LineSpacingMultiple = false;
        last.SpaceBeforeTwips = 0;
        last.SpaceAfterTwips = 0;
        document.AddParagraph("Body").AddNoteReference(note);
        for (int pass = 0; pass < 3; pass++) {
            document = RtfDocument.Read(document.ToRtf()).Document;
            RtfNote reopened = Assert.Single(document.Notes);
            Assert.Equal(2, reopened.Paragraphs.Count);
            AssertTerminalState(reopened.Paragraphs[1]);
        }
    }

    private static void AssertTerminalState(RtfParagraph paragraph) {
        Assert.Empty(paragraph.Inlines);
        Assert.Equal(-360, paragraph.LineSpacingTwips);
        Assert.Equal(false, paragraph.LineSpacingMultiple);
        Assert.Equal(0, paragraph.SpaceBeforeTwips);
        Assert.Equal(0, paragraph.SpaceAfterTwips);
    }

    // Exact RTF content of the already-qualified native placement-explicit-ftnbj
    // control, without its final file newline. No new native oracle is required.
    private const string NativeSource = @"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Arial;}}{\stylesheet{\s0\ql\f0\fs24 Normal;}}\paperw12240\paperh15840\margl1440\margr1440\margt1440\margb1440\deftab720\ftnbj\pard\plain\s0\f0\fs24\sb0\sa0 NOTE-BODY-BEFORE {\super\chftn}{\footnote\pard\plain\s0\f0\fs24\sb0\sa0\sl360\slmult0 {\super\chftn} NOTE-MIN-A {\fs48 NOTE-MIN-TALL} END-NOTE-MIN\line NOTE-MIN-B\par\pard\plain\s0\f0\fs24\sb0\sa0\sl-360\slmult0 NOTE-EXACT-A {\fs48 NOTE-EXACT-TALL} END-NOTE-EXACT\line NOTE-EXACT-B\par} NOTE-BODY-AFTER\par}";
}
