using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void NoteTextParagraphsPreserveLogicalBoundariesAndVisibleContent() {
        using WordDocument source = WordDocument.Create();
        WordParagraph owner = source.AddParagraph("owner");
        owner.AddBookmark("LocalAnchor");
        owner.AddFootNote("footnote");
        owner.AddEndNote("endnote");
        var main = source.OpenXmlDocument.MainDocumentPart!;
        FootnoteEndnoteType[] bodies = {
            Assert.Single(main.FootnotesPart!.Footnotes!.Elements<Footnote>(), note => note.Type == null),
            Assert.Single(main.EndnotesPart!.Endnotes!.Elements<Endnote>(), note => note.Type == null)
        };
        foreach (FootnoteEndnoteType body in bodies) {
            body.RemoveAllChildren<Paragraph>();
            var rich = new Paragraph(
                new Run(new Text("left ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new RunProperties(new Bold()), new Text("bold")),
                new Hyperlink(new Run(new Text(" link") { Space = SpaceProcessingModeValues.Preserve })) { Anchor = "LocalAnchor" },
                new SimpleField(new Run(new Text(" cached") { Space = SpaceProcessingModeValues.Preserve })) { Instruction = " FILENAME " },
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" SECRET_INSTRUCTION ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(" visible") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new TabChar(), new Text(" end") { Space = SpaceProcessingModeValues.Preserve }));
            rich.PrependChild(body is Footnote ? new Run(new FootnoteReferenceMark()) : new Run(new EndnoteReferenceMark()));
            body.Append(rich, new Paragraph(),
                new Paragraph(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }, new TabChar(),
                    new Text(" ") { Space = SpaceProcessingModeValues.Preserve })),
                new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("nested table text")))))));
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        main = document.OpenXmlDocument.MainDocumentPart!;
        WordFootNote footnote = Assert.Single(document.FootNotes);
        WordEndNote endnote = Assert.Single(document.EndNotes);
        string[] expected = { "left bold link cached visible\t end", "", " \t " };
        string[] xml = { main.Document.OuterXml, main.FootnotesPart!.Footnotes!.OuterXml, main.EndnotesPart!.Endnotes!.OuterXml };
        Assert.Equal(expected, footnote.TextParagraphs);
        Assert.Equal(expected, endnote.TextParagraphs);
        Assert.Equal(expected, Assert.Single(document.Sections[0].FootNotes).TextParagraphs);
        Assert.Equal(expected, Assert.Single(document.Sections[0].EndNotes).TextParagraphs);
        Assert.Equal(xml, new[] { main.Document.OuterXml, main.FootnotesPart.Footnotes.OuterXml, main.EndnotesPart.Endnotes.OuterXml });

        Footnote footBody = Assert.Single(main.FootnotesPart.Footnotes.Elements<Footnote>(), note => note.Id?.Value == footnote.ReferenceId);
        Endnote endBody = Assert.Single(main.EndnotesPart.Endnotes.Elements<Endnote>(), note => note.Id?.Value == endnote.ReferenceId);
        footBody.Type = FootnoteEndnoteValues.Separator;
        endBody.Type = FootnoteEndnoteValues.Separator;
        Assert.Null(footnote.TextParagraphs);
        Assert.Null(endnote.TextParagraphs);
        footBody.Remove();
        endBody.Remove();
        Assert.Null(footnote.TextParagraphs);
        Assert.Null(endnote.TextParagraphs);
    }

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
        noteParagraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(36D);
        noteParagraph.IndentationHanging = global::OfficeIMO.OfficeLength.Points(18D);
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
