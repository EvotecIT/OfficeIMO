using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdtNotesTests {
    [Fact]
    public void NoteInRepeatedTableCellChangesOnlySelectedLogicalCell() {
        OdtDocument document = OdtDocument.Create();
        OdtTable table = document.AddTable(1, 1, "RepeatedNotes");
        table.Cell(0, 0).Text = "Anchor";
        XElement row = table.Element.Elements(OdfNamespaces.Table + "table-row").Single();
        row.SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 3);
        row.Elements(OdfNamespaces.Table + "table-cell").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 3);
        document.Package.MarkXmlDirty("content.xml");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        byte[] beforeRead = reopened.ToBytes();
        Assert.Equal("Anchor", reopened.Tables.Single().Rows[1].Cells[1].Paragraphs[0].Text);
        Assert.Equal(beforeRead, reopened.ToBytes());
        reopened.Tables.Single().Rows[1].Cells[1].Paragraphs[0].AddFootnote("Selected note");

        OdtDocument result = OdtDocument.Load(new MemoryStream(reopened.ToBytes()));
        OdtTable resultTable = result.Tables.Single();
        for (int rowIndex = 0; rowIndex < 3; rowIndex++) {
            for (int columnIndex = 0; columnIndex < 3; columnIndex++) {
                OdtParagraph paragraph = resultTable.Rows[rowIndex].Cells[columnIndex].Paragraphs[0];
                if (rowIndex == 1 && columnIndex == 1)
                    Assert.Equal("Selected note", Assert.Single(paragraph.Notes).Paragraphs[0].Text);
                else Assert.Empty(paragraph.Notes);
            }
        }
        Assert.True(result.Validate().IsValid);
    }

    [Fact]
    public void RejectedNoteInRepeatedTableCellDoesNotSplitTable() {
        OdtDocument document = OdtDocument.Create();
        OdtTable table = document.AddTable(1, 1, "ConfiguredNotes");
        table.Element.Descendants(OdfNamespaces.Table + "table-cell").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 3);
        document.Package.MarkXmlDirty("content.xml");
        document.Package.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
            new XElement(OdfNamespaces.Text + "notes-configuration",
                new XAttribute(OdfNamespaces.Text + "note-class", "footnote"),
                new XAttribute(OdfNamespaces.Text + "start-value", "5")));
        document.Package.MarkXmlDirty("styles.xml");
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        byte[] before = reopened.ToBytes();

        Assert.Throws<NotSupportedException>(() =>
            reopened.Tables.Single().Rows[0].Cells[1].Paragraphs[0].AddFootnote("Rejected"));
        Assert.Equal(before, reopened.ToBytes());
    }

    [Fact]
    public void SupportedNotesAreReportedAsEditable() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Anchor").AddFootnote("Body");

        OdfFeatureReport report = source.InspectFeatures();
        Assert.Contains(report.Findings, finding => finding.Name == "text-notes" &&
            finding.Support == OdfFeatureSupport.Editable && finding.Count == 1);
    }

    [Fact]
    public void NoteInsideTrackedDeletionIsNotReportedAsEditable() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph deleted = source.AddParagraph("Anchor");
        deleted.AddFootnote("Deleted note");
        source.DeleteParagraphTracked(deleted, "Editor");

        OdfFeatureReport report = source.InspectFeatures();
        Assert.DoesNotContain(report.Findings, finding => finding.Name == "text-notes" &&
            finding.Support == OdfFeatureSupport.Editable);
        Assert.Contains(report.Findings, finding => finding.Name == "text-notes" &&
            finding.Support == OdfFeatureSupport.Inspected && finding.Count == 1);
    }

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
    public void ConfiguredNumberingRejectsNewNotesWithoutChangingExistingCitations() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph anchor = source.AddParagraph("Anchor");
        anchor.AddFootnote("Existing");
        var office = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        var text = (System.Xml.Linq.XNamespace)"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        source.Package.GetXml("styles.xml").Root!.Element(office + "styles")!.Add(
            new System.Xml.Linq.XElement(text + "notes-configuration",
                new System.Xml.Linq.XAttribute(text + "note-class", "footnote"),
                new System.Xml.Linq.XAttribute(text + "start-value", "5")));
        source.Package.MarkXmlDirty("styles.xml");
        byte[] before = source.ToBytes();

        Assert.Throws<NotSupportedException>(() => anchor.AddFootnote("New"));
        Assert.Equal(before, source.ToBytes());
        Assert.Equal("1", Assert.Single(anchor.Notes).Citation);
        anchor.AddEndnote("An unrelated note kind");
    }

    [Fact]
    public void ManyAppendedNotesAndEarlierInsertionKeepPackageOrder() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph first = source.AddParagraph("First");
        for (int index = 0; index < 300; index++)
            source.AddParagraph("Anchor " + index).AddFootnote("Note " + index);
        first.AddFootnote("Inserted first");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        string[] citations = reopened.Paragraphs.SelectMany(paragraph => paragraph.Notes)
            .Select(note => note.Citation).ToArray();
        Assert.Equal(301, citations.Length);
        Assert.Equal(Enumerable.Range(1, 301).Select(number => number.ToString()).ToArray(), citations);
        Assert.Equal(301, reopened.Paragraphs.SelectMany(paragraph => paragraph.Notes)
            .Select(note => note.Id).Distinct().Count());
    }

    [Fact]
    public void ReplacingParagraphTextRebuildsNoteOrdinals() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph first = source.AddParagraph("First");
        OdtParagraph second = source.AddParagraph("Second");
        OdtParagraph third = source.AddParagraph("Third");
        first.AddFootnote("Removed");
        second.AddFootnote("Kept");

        first.Text = "Replaced";
        third.AddFootnote("Appended");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal(new[] { "1", "2" }, reopened.Paragraphs.SelectMany(paragraph => paragraph.Notes)
            .Select(note => note.Citation));
    }

    [Fact]
    public void ReplacingParagraphTextAfterLoadRenumbersGeneratedCitations() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("First").AddFootnote("Removed");
        source.AddParagraph("Second").AddFootnote("Kept");

        OdtDocument loaded = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        loaded.Paragraphs[0].Text = "Replaced";
        Assert.Equal("1", Assert.Single(loaded.Paragraphs[1].Notes).Citation);
        loaded.AddParagraph("Third").AddFootnote("Appended");

        OdtDocument reopened = OdtDocument.Load(new MemoryStream(loaded.ToBytes()));
        Assert.Equal(new[] { "1", "2" }, reopened.Paragraphs.SelectMany(paragraph => paragraph.Notes)
            .Select(note => note.Citation));
    }

    [Fact]
    public void TrackedDeletionReservesItsNoteIdUntilAccepted() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph first = source.AddParagraph("First");
        string originalId = first.AddFootnote("Temporarily deleted").Id!;
        OdtTrackedChange deletion = source.DeleteParagraphTracked(first, "Editor");
        OdtDocument loaded = OdtDocument.Load(new MemoryStream(source.ToBytes()));
        string newId = loaded.AddParagraph("Second").AddFootnote("New note").Id!;

        Assert.NotEqual(originalId, newId);
        Assert.True(loaded.RejectTrackedChange(deletion.Id));
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(loaded.ToBytes()));
        OdtNote[] notes = reopened.Paragraphs.SelectMany(paragraph => paragraph.Notes).ToArray();
        Assert.Equal(2, notes.Length);
        Assert.Equal(2, notes.Select(note => note.Id).Distinct().Count());
        Assert.Equal(new[] { "1", "2" }, notes.Select(note => note.Citation));
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
