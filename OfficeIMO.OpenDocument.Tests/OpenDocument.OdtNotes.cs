using System;
using System.IO;
using System.Linq;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdtNotesTests {
    [Fact]
    public void InsertingEarlierNoteRenumbersGeneratedCitationsInDocumentOrder() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph earlier = source.AddParagraph("Earlier");
        OdtParagraph later = source.AddParagraph("Later");
        later.AddFootnote("Later note");
        earlier.AddFootnote("Earlier note");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal(new[] { "1", "2" }, reopened.Paragraphs.SelectMany(paragraph => paragraph.Notes)
            .Select(note => note.Citation));
        Assert.True(reopened.Validate().IsValid);
    }

    [Fact]
    public void InsertingBodyNotePersistsRenumberedHeaderCitation() {
        OdtDocument created = OdtDocument.Create();
        created.PageLayout.Header.AddParagraph("Header").AddFootnote("Header note");
        OdtDocument loaded = OdtDocument.Load(new MemoryStream(created.ToBytes()));
        loaded.AddParagraph("Body").AddFootnote("Body note");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(loaded.ToBytes()));
        Assert.Equal("1", Assert.Single(Assert.Single(reopened.Paragraphs).Notes).Citation);
        Assert.Equal("2", Assert.Single(Assert.Single(reopened.PageLayout.Header.Paragraphs).Notes).Citation);
    }

    [Fact]
    public void InsertingNoteLeavesCustomCitationUntouched() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph earlier = source.AddParagraph("Earlier");
        OdtParagraph later = source.AddParagraph("Later");
        later.AddEndnote("Later note");
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        source.Package.GetXml("content.xml").Descendants(text + "note-citation").Single().Value = "*";
        source.Package.MarkXmlDirty("content.xml");
        earlier.AddEndnote("Earlier note");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal(new[] { "1", "*" }, reopened.Paragraphs.SelectMany(paragraph => paragraph.Notes)
            .Select(note => note.Citation));
    }

    [Fact]
    public void NoteBodyInlineCollectionsBelongToNoteParagraphOnly() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph anchor = source.AddParagraph("Anchor");
        OdtParagraph body = anchor.AddFootnote("Body").Paragraphs[0];
        body.AddSpan(" emphasis").Bold = true;
        body.AddHyperlink(" link", "https://example.com");
        body.AddImage(Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="),
            "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        OdtParagraph reopenedAnchor = Assert.Single(reopened.Paragraphs);
        OdtParagraph reopenedBody = Assert.Single(Assert.Single(reopenedAnchor.Notes).Paragraphs);
        Assert.Empty(reopenedAnchor.Spans);
        Assert.Empty(reopenedAnchor.Hyperlinks);
        Assert.Empty(reopenedAnchor.Images);
        Assert.Single(reopenedBody.Spans);
        Assert.Single(reopenedBody.Hyperlinks);
        Assert.Single(reopenedBody.Images);
    }
}
