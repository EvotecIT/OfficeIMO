using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Syntax;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfNativeListAndNoteGrammarTests {
    [Theory]
    [InlineData(RtfNoteKind.Footnote)]
    [InlineData(RtfNoteKind.Endnote)]
    public void NativeNotesHaveOneInternalReferenceBeforeTheFullAuthoredText(RtfNoteKind kind) {
        RtfDocument document = RtfDocument.Create();
        var note = new RtfNote(kind);
        note.AddParagraph("Full note text");
        document.AddParagraph("Body").AddNoteReference(note);
        for (int pass = 0; pass < 3; pass++) {
            document = RtfDocument.Read(document.ToRtf()).Document;
            RtfNote reopened = Assert.Single(document.Notes);
            Assert.Equal("Full note text", reopened.ToPlainText());
            RtfGeneratedText marker = Assert.IsType<RtfGeneratedText>(reopened.Paragraphs[0].Inlines[0]);
            Assert.Equal(RtfGeneratedTextKind.NoteReference, marker.Kind);
            Assert.Null(marker.Note);
            Assert.Single(reopened.Paragraphs[0].Inlines.OfType<RtfGeneratedText>());
        }
        Assert.Single(note.Paragraphs[0].Inlines);
    }

    [Fact]
    public void NativeDocumentsDeclareBothNoteKindsWithoutConvertingThem() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Body").AddFootnote("1", "Footnote");
        document.AddParagraph("Body").AddEndnote("E", "Endnote");
        RtfReadResult result = RtfDocument.Read(document.ToRtf());
        Assert.Equal(2, Assert.Single(result.SyntaxTree.Root.Children.OfType<RtfControlWord>(), item => item.Name == "fet").Parameter);
        Assert.Equal(new[] { RtfNoteKind.Footnote, RtfNoteKind.Endnote }, result.Document.Notes.Select(note => note.Kind));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NoteParagraphsTerminateBetweenItemsAndRetainAnExplicitEmptyLastItem(bool emptyLast) {
        var note = new RtfNote(RtfNoteKind.Footnote);
        note.AddParagraph("First").SetList();
        note.AddParagraph(emptyLast ? null : "Second").SetList();
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Body").AddNoteReference(note);
        RtfReadResult result = RtfDocument.Read(document.ToRtf());
        RtfGroup noteGroup = Assert.Single(result.SyntaxTree.Root.Children.OfType<RtfGroup>(), item => item.Destination == "footnote");
        Assert.Single(noteGroup.Children.OfType<RtfControlWord>(), item => item.Name == "par");
        Assert.Equal(new[] { "First", emptyLast ? "" : "Second" }, result.Document.Notes[0].Paragraphs.Select(paragraph => paragraph.ToPlainText()));
    }

    [Fact]
    public void ModernListBindingsDoNotWriteAnUnscopedLegacyDestination() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Bullet").SetList();
        document.AddParagraph("Anchor").AddFootnote("1", "Note bullet").Note!.Paragraphs[0].SetList(3);
        RtfReadResult reopened = RtfDocument.Read(document.ToRtf());
        Assert.DoesNotContain(reopened.SyntaxTree.Root.Children.OfType<RtfControlWord>(), item => item.Name == "pn");
        RtfGroup noteGroup = Assert.Single(reopened.SyntaxTree.Root.Children.OfType<RtfGroup>(), item => item.Destination == "footnote");
        Assert.DoesNotContain(noteGroup.Children.OfType<RtfControlWord>(), item => item.Name == "pn");
        Assert.Equal(RtfListKind.Bullet, reopened.Document.ResolveListFormatting(reopened.Document.Paragraphs[0])!.Level.Kind);
        Assert.Equal(RtfListKind.Bullet, reopened.Document.ResolveListFormatting(reopened.Document.Notes[0].Paragraphs[0])!.Level.Kind);
        Assert.Equal("Anchor1", reopened.Document.Paragraphs[1].ToPlainText());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void ListTemplatesUseCountedUnicodeCharactersForNativeReaders(int skipCount) {
        const string template = "•—\u200E";
        RtfDocument document = RtfDocument.Create();
        document.Settings.UnicodeSkipCount = skipCount;
        document.AddListDefinition(10).AddLevel(RtfListKind.Bullet).Text = template;
        document.AddListOverride(1, 10);
        document.AddParagraph("Bullet").SetList();
        var result = RtfDocument.Read(document.ToRtf());
        RtfGroup table = Assert.Single(result.SyntaxTree.Root.Children.OfType<RtfGroup>(), item => item.Destination == "listtable");
        RtfGroup level = Assert.Single(Assert.Single(table.Children.OfType<RtfGroup>()).Children.OfType<RtfGroup>(), item => item.Destination == "listlevel");
        RtfGroup text = Assert.Single(level.Children.OfType<RtfGroup>(), item => item.Destination == "leveltext");
        Assert.Equal(new[] { 8226, 8212, 8206 }, text.Children.OfType<RtfControlWord>().Where(item => item.Name == "u").Select(item => item.Parameter!.Value));
        Assert.Equal(template, result.Document.ListDefinitions[0].Levels[0].Text);
    }

    [Fact]
    public void StandardEndnoteDestinationRetainsItsKindAndAutomaticReference() {
        const string input = @"{\rtf1\ansi Body\chftn{\footnote\ftnalt\pard Endnote text\par}\par}";
        RtfReadResult read = RtfDocument.Read(input);
        Assert.Equal(RtfNoteKind.Endnote, Assert.Single(read.Document.Notes).Kind);
        Assert.Equal(input, read.ToRtfLossless());
        string output = read.Document.ToRtf();
        Assert.Contains(@"{\footnote\ftnalt", output, StringComparison.Ordinal);
        RtfDocument reopened = RtfDocument.Read(output).Document;
        Assert.Equal(RtfNoteKind.Endnote, Assert.Single(reopened.Notes).Kind);
        Assert.Same(reopened.Notes[0], Assert.Single(reopened.Paragraphs[0].Inlines.OfType<RtfGeneratedText>()).Note);
        Assert.Equal("Endnote text", reopened.Notes[0].ToPlainText());
    }
}
