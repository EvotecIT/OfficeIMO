using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Testing;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordOdtNotesConversionTests {
    [Fact]
    public void ExistingWordNoteFixtureKeepsBothNoteKindsThroughOdt() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "DocumentWithFootNotes.docx");
        using WordDocument source = WordDocument.Load(path);
        OdfConversionResult<OdtDocument> conversion = source.ToOpenDocumentResult();
        Assert.True(conversion.Value.Validate().IsValid);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "footnotes" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 3);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "endnotes" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 2);

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(conversion.Value.ToBytes()));
        OdtNote[] notes = reopened.Paragraphs.SelectMany(paragraph => paragraph.Notes).ToArray();
        Assert.Equal(3, notes.Count(note => note.Kind == OdtNoteKind.Footnote));
        Assert.Equal(2, notes.Count(note => note.Kind == OdtNoteKind.Endnote));
        Assert.Contains(notes, note => note.Paragraphs.Any(paragraph => paragraph.Text.Contains("first footnote", StringComparison.Ordinal)));
        Assert.Contains(notes, note => note.Paragraphs.Any(paragraph => paragraph.Text.Contains("1st end note", StringComparison.Ordinal)));
    }

    [Fact]
    public void WordFootnotesAndEndnotesSurviveOdtAndDocxPackages() {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Before ");
        paragraph.AddFootNote("Footnote text");
        paragraph.AddText(" between ");
        paragraph.AddEndNote("Endnote text");
        paragraph.AddText(" after");

        OdfConversionResult<OdtDocument> toOdt = source.ToOpenDocumentResult();
        Assert.True(toOdt.Value.Validate().IsValid);
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "footnotes" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 1);
        Assert.Contains(toOdt.Report.Mappings, mapping => mapping.Feature == "endnotes" &&
            mapping.Status == OdfConversionMappingStatus.Converted && mapping.Count == 1);
        Assert.DoesNotContain(toOdt.Report.Mappings, mapping => mapping.Feature == "source-footnotes" ||
            mapping.Feature == "source-endnotes");

        using var package = new MemoryStream(toOdt.Value.ToBytes());
        OdtDocument reopened = OdtDocument.Load(package);
        OdtParagraph odtParagraph = Assert.Single(reopened.ContentBlocks).Paragraph!;
        Assert.Equal("Before  between  after", odtParagraph.Text);
        Assert.Equal(new OdtNoteKind?[] { OdtNoteKind.Footnote, OdtNoteKind.Endnote },
            odtParagraph.Notes.Select(note => note.Kind).ToArray());
        Assert.Equal("Footnote text", odtParagraph.Notes[0].Paragraphs[0].Text);
        Assert.Equal("Endnote text", odtParagraph.Notes[1].Paragraphs[0].Text);
        Assert.Equal(new[] { OdtInlineNodeKind.Span, OdtInlineNodeKind.Note, OdtInlineNodeKind.Span,
            OdtInlineNodeKind.Note, OdtInlineNodeKind.Span }, odtParagraph.InlineNodes.Select(node => node.Kind));
        Assert.Single(reopened.Paragraphs);

        OdfConversionResult<WordDocument> toWord = reopened.ToWordDocumentResult();
        Assert.DoesNotContain(toWord.Report.Mappings, mapping => mapping.Feature == "source-text-notes");
        using WordDocument word = toWord.Value;
        Assert.Empty(word.ValidateDocument());
        using WordDocument saved = WordDocument.Load(new MemoryStream(word.ToBytes()));
        WordParagraphSnapshot snapshot = Assert.Single(saved.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>());
        Assert.Equal("Footnote text", Assert.Single(snapshot.Runs, run => run.Footnote != null)
            .Footnote!.Paragraphs[0].Text);
        Assert.Equal("Endnote text", Assert.Single(snapshot.Runs, run => run.Endnote != null)
            .Endnote!.Paragraphs[0].Text);
    }

    [Fact]
    public void OdtNoteBodiesDoNotBecomeBodyParagraphsAndRichBodyLossIsExplicit() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddParagraph("Body");
        OdtNote note = paragraph.AddFootnote("First note paragraph");
        note.AddParagraph("Second note paragraph");
        paragraph.AddEndnote("End note");
        Assert.Single(source.Paragraphs);
        Assert.Single(source.ContentBlocks);

        using var package = new MemoryStream(source.ToBytes());
        OdtDocument reopened = OdtDocument.Load(package);
        Assert.Equal(2, Assert.Single(reopened.Paragraphs).Notes.Count);
        OdfConversionResult<WordDocument> conversion = reopened.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-formatting" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Throws<OdfConversionLossException>(() => reopened.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void CustomOdtCitationsAndUnmodeledNoteMediaRemainExplicitLoss() {
        OdtDocument source = OdtDocument.Create();
        OdtNote note = source.AddParagraph("Cited text").AddFootnote("Source text");
        note.Paragraphs[0].AddImage(Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="),
            "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
        byte[] modified = OdfTestPackageRewriter.Rewrite(source.ToBytes(), (name, bytes) => {
            if (name != "content.xml") return bytes;
            XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            xml.Descendants(text + "note-citation").Single().Value = "*";
            return Encoding.UTF8.GetBytes(xml.ToString());
        });
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(modified));
        OdfConversionResult<WordDocument> conversion = reopened.ToWordDocumentResult();
        using WordDocument word = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-citations" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "note-body-content" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => reopened.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }
}
