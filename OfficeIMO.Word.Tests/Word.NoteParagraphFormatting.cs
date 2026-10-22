using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NoteAuthoringRetainsDefaultParagraphStyleAndReferenceMarker(bool endnote) {
        using WordDocument document = WordDocument.Create();
        WordParagraph body = document.AddParagraph("Reference");
        if (endnote) body.AddEndNote("Note content"); else body.AddFootNote("Note content");

        AssertDefaultAuthoredNote(document, endnote);
        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        AssertDefaultAuthoredNote(reopened, endnote);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void NoteCreationPreservesAuthoredParagraphProperties(bool endnote, bool explicitStyle) {
        using WordDocument document = WordDocument.Create();
        WordParagraph body = document.AddParagraph("Reference");
        var noteParagraph = new WordParagraph(document) { Text = "Note content" };
        if (explicitStyle) noteParagraph.Style = WordParagraphStyles.Heading2;
        noteParagraph.ParagraphAlignment = WordParagraphAlignment.Right;
        noteParagraph.IndentationBefore = 720;
        noteParagraph.IndentationHanging = 360;
        noteParagraph.KeepWithNext = true;
        string? expectedStyle = noteParagraph.StyleId;

        if (endnote) WordEndNote.AddEndNote(document, body, noteParagraph);
        else WordFootNote.AddFootNote(document, body, noteParagraph);

        AssertAuthoredNoteProperties(document, endnote, explicitStyle ? expectedStyle! : endnote ? "EndnoteText" : "FootnoteText");
        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        AssertAuthoredNoteProperties(reopened, endnote, explicitStyle ? expectedStyle! : endnote ? "EndnoteText" : "FootnoteText");
        Assert.Empty(reopened.ValidateDocument());
    }

    private static void AssertDefaultAuthoredNote(WordDocument document, bool endnote) {
        Paragraph paragraph = GetAuthoredNoteParagraph(document, endnote);
        Assert.Equal(endnote ? "EndnoteText" : "FootnoteText", paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.Equal("Note content", paragraph.InnerText);
        AssertNoteReferenceMark(paragraph, endnote);
    }

    [Theory]
    [InlineData(false, "field-only")]
    [InlineData(false, "hyperlink-only")]
    [InlineData(false, "hyperlink-before-text")]
    [InlineData(true, "field-only")]
    [InlineData(true, "hyperlink-only")]
    [InlineData(true, "hyperlink-before-text")]
    public void NoteReferenceMarkerPrecedesAllAuthoredInlineContent(bool endnote, string shape) {
        using WordDocument document = WordDocument.Create();
        WordParagraph body = document.AddParagraph("Reference");
        body.AddBookmark("LocalAnchor");
        var noteParagraph = new WordParagraph(document, newRun: false);
        if (shape == "field-only") {
            noteParagraph._paragraph.Append(new SimpleField(new Run(new Text("Inline content"))) { Instruction = " PAGE " });
        } else {
            noteParagraph._paragraph.Append(new Hyperlink(new Run(new Text("Inline content"))) { Anchor = "LocalAnchor" });
            if (shape == "hyperlink-before-text") noteParagraph._paragraph.Append(new Run(new Text(" tail") { Space = SpaceProcessingModeValues.Preserve }));
        }
        string[] originalContent = noteParagraph._paragraph.ChildElements.Skip(1).Select(child => child.OuterXml).ToArray();
        if (endnote) WordEndNote.AddEndNote(document, body, noteParagraph);
        else WordFootNote.AddFootNote(document, body, noteParagraph);

        AssertReferenceOrder(noteParagraph._paragraph, endnote, originalContent);
        Assert.Empty(document.ValidateDocument());
        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        var main = reopened._wordprocessingDocument.MainDocumentPart!;
        OpenXmlElement root = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        Paragraph actual = Assert.Single(root.Descendants<Paragraph>(), paragraph => paragraph.InnerText.StartsWith("Inline content", StringComparison.Ordinal));
        AssertReferenceOrder(actual, endnote, originalContent);
        Assert.Empty(reopened.ValidateDocument());
    }

    private static void AssertReferenceOrder(Paragraph paragraph, bool endnote, string[] originalContent) {
        Assert.IsType<ParagraphProperties>(paragraph.ChildElements[0]);
        Run reference = Assert.IsType<Run>(paragraph.ChildElements[1]);
        AssertNoteReferenceMark(paragraph, endnote);
        if (endnote) Assert.Single(reference.Elements<EndnoteReferenceMark>());
        else Assert.Single(reference.Elements<FootnoteReferenceMark>());
        Assert.Equal(originalContent, paragraph.ChildElements.Skip(2).Select(child => child.OuterXml));
    }

    private static void AssertAuthoredNoteProperties(WordDocument document, bool endnote, string style) {
        Paragraph paragraph = GetAuthoredNoteParagraph(document, endnote);
        Assert.Equal(style, paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        Assert.NotNull(paragraph.ParagraphProperties.Justification);
        Assert.NotNull(paragraph.ParagraphProperties.Indentation);
        Assert.NotNull(paragraph.ParagraphProperties.KeepNext);
        Assert.Equal(JustificationValues.Right, paragraph.ParagraphProperties.Justification!.Val!.Value);
        Assert.Equal("720", paragraph.ParagraphProperties.Indentation!.Left!.Value);
        Assert.Equal("360", paragraph.ParagraphProperties.Indentation.Hanging!.Value);
        Assert.True(paragraph.ParagraphProperties.KeepNext!.Val?.Value ?? true);
        AssertNoteReferenceMark(paragraph, endnote);
    }

    private static Paragraph GetAuthoredNoteParagraph(WordDocument document, bool endnote) {
        var main = document._wordprocessingDocument.MainDocumentPart!;
        OpenXmlElement root = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        return Assert.Single(root.Descendants<Paragraph>(), paragraph => paragraph.InnerText == "Note content");
    }

    private static void AssertNoteReferenceMark(Paragraph paragraph, bool endnote) {
        if (endnote) Assert.Single(paragraph.Descendants<EndnoteReferenceMark>());
        else Assert.Single(paragraph.Descendants<FootnoteReferenceMark>());
    }
}
