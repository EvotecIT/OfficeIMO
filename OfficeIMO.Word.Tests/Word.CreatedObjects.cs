using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordCreatedObjectTests {
    [Fact]
    public void ReturnedHeadersAndFootersRemainAttachedToTheirSection() {
        using var document = WordDocument.Create();
        var first = document.AddHeadersAndFooters();
        first.Header.AddParagraph("First header");
        first.Footer.AddParagraph("First footer");
        var section = document.AddSection();
        var second = section.AddHeadersAndFooters();
        second.Header.AddParagraph("Second header");
        second.Footer.AddParagraph("Second footer");
        Assert.Same(first.Header, document.Sections[0].Header.Default);
        Assert.Same(second.Footer, section.Footer.Default);
        Assert.Same(second.Header, section.AddHeadersAndFooters().Header);
        using var loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Contains(loaded.Sections[0].Header.Default!.Paragraphs, paragraph => paragraph.Text == "First header");
        Assert.Contains(loaded.Sections[1].Footer.Default!.Paragraphs, paragraph => paragraph.Text == "Second footer");
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(loaded._wordprocessingDocument));
    }

    [Fact]
    public void FieldsRetainTheirOwnNativeRepresentationWhenTheParagraphAddsMoreFields() {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        WordField first = paragraph.AddField(WordFieldType.Page);
        WordField second = paragraph.AddField(WordFieldType.NumPages, advanced: true);
        Assert.Equal(WordFieldType.Page, first.FieldType);
        Assert.Equal(WordFieldType.NumPages, second.FieldType);
        first.Text = "12";
        first.LockField = true;
        second.Text = "24";
        Assert.Equal("12", first.Text);
        Assert.Equal("24", second.Text);
        Assert.Same(paragraph, first.Paragraph);
        Assert.Same(paragraph, second.Paragraph);

        using var loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Contains(loaded.Fields, field => field.FieldType == WordFieldType.Page && field.LockField && field.Text == "12");
        Assert.Contains(loaded.Fields, field => field.FieldType == WordFieldType.NumPages && field.Text == "24");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CommentsAttachToEmptyParagraphsAndHyperlinkOnlyParagraphs(bool linked) {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        if (linked) paragraph.AddHyperLink("Linked text", new Uri("https://example.com/"));
        WordComment comment = paragraph.AddComment("Reviewer", "RV", "Initial");
        comment.Text = "Edited";
        WordComment tableComment = document.AddTable(1, 1).AddComment("Reviewer", "RV", "Table");
        tableComment.Text = "Edited table";
        using var loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(new[] { "Edited", "Edited table" }, loaded.Comments.Select(item => item.Text));
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(loaded._wordprocessingDocument));
    }

    [Fact]
    public void ReturnedNotesAddressTheNewReferenceAndEditableBody() {
        using var document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Text");
        WordFootNote first = paragraph.AddFootNote("First");
        WordFootNote second = paragraph.AddFootNote("Second");
        WordEndNote end = paragraph.AddEndNote("End");
        Assert.NotEqual(first.ReferenceId, second.ReferenceId);
        first.Paragraphs![0].Text = "Updated first";
        second.Paragraphs![0].Text = "Updated second";
        end.Paragraphs![0].Text = "Updated end";
        Assert.Equal(first.ReferenceId, first.Paragraph.FootNote!.ReferenceId);
        Assert.Equal(second.ReferenceId, second.Paragraph.FootNote!.ReferenceId);
        using var loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Contains(loaded.FootNotes, note => note.Paragraphs!.Any(body => body.Text == "Updated first"));
        Assert.Contains(loaded.FootNotes, note => note.Paragraphs!.Any(body => body.Text == "Updated second"));
        Assert.Contains(loaded.EndNotes, note => note.Paragraphs!.Any(body => body.Text == "Updated end"));
    }

    [Fact]
    public void BookmarksLinksAndEquationsReturnAttachedEditableObjects() {
        using var document = WordDocument.Create();
        WordBookmark bookmark = document.AddBookmark("Original");
        bookmark.Name = "Destination";
        WordHyperLink link = document.AddHyperLink("Go", "Original");
        link.Anchor = "Destination";
        WordEquation equation = document.AddEquation("<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>");
        equation.Text = "y";
        using var loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Contains(loaded.Bookmarks, item => item.Name == "Destination");
        Assert.Contains(loaded.HyperLinks, item => item.Anchor == "Destination");
        Assert.Equal("y", loaded.Equations.Single().Text);
    }
}
