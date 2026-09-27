using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests;

public sealed class WordImageVmlOwnershipTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DrawingMlClone_ResolvesBytesInAnotherStoryOrDocument(bool otherDocument) {
        using var sourceDocument = WordDocument.Create();
        using var destination = WordDocument.Create();
        byte[] expected = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
        var source = sourceDocument.AddParagraph().AddImage(new MemoryStream(expected), "shared.gif", 20, 20).Image!;
        var targetDocument = otherDocument ? destination : sourceDocument;
        var target = targetDocument.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph();
        var clone = source.Clone(target);
        source.Remove();
        Assert.Equal(expected, clone.ToBytes());
        using var bytes = targetDocument.ToStream();
        using var reopened = WordDocument.Load(bytes);
        Assert.Equal(expected, reopened.Sections[0].Header.Default!.Paragraphs.Single(item => item.IsImage).Image!.ToBytes());
    }
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void VmlClone_PreservesItsShapeAndRelationshipInDestinationStory(bool external, bool otherDocument) {
        using var document = WordDocument.Create();
        using var destination = WordDocument.Create();
        byte[] expected = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
        var paragraph = document.AddParagraph("Image");
        var owner = document.MainDocumentPartRoot;
        string id;
        if (external) id = owner.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", new Uri("https://example.test/shared.gif")).Id;
        else { var part = owner.AddImagePart(ImagePartType.Gif); part.FeedData(new MemoryStream(expected)); id = owner.GetIdOfPart(part); }
        paragraph._run!.Append(new W.Picture(new V.Shape(new V.ImageData { RelationshipId = id }) { Id = "Original", Style = "width:15pt;height:15pt" }));
        var targetDocument = otherDocument ? destination : document;
        var target = targetDocument.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph();
        var clone = paragraph.Image!.Clone(target);
        Assert.Equal(20, clone.Width);
        Assert.Single(target._paragraph.Descendants<V.Shape>());
        Assert.Empty(target._paragraph.Descendants<W.Drawing>());
        if (external) Assert.Equal("https://example.test/shared.gif", clone.ExternalUri!.AbsoluteUri);
        else Assert.Equal(expected, clone.ToBytes());
        paragraph.Image!.Remove();
        if (external) Assert.NotNull(clone.ExternalUri); else Assert.Equal(expected, clone.ToBytes());
        using var bytes = targetDocument.ToStream();
        using var reopened = WordDocument.Load(bytes);
        var persisted = reopened.Sections[0].Header.Default!.Paragraphs.Single(item => item.IsImage).Image!;
        if (external) Assert.NotNull(persisted.ExternalUri); else Assert.Equal(expected, persisted.ToBytes());
    }
    [Fact]
    public void ImageRemoval_PreservesClonedDrawingMlOccurrence() {
        using var document = WordDocument.Create();
        byte[] expected = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
        var source = document.AddParagraph().AddImage(new MemoryStream(expected), "shared.gif", 20, 20).Image!;
        var clone = source.Clone(document.AddParagraph());
        source.Remove();
        Assert.Equal(expected, clone.ToBytes());
        Assert.NotNull(document.Paragraphs[1].Image);
        clone.Remove();
        Assert.Null(document.Paragraphs[1].Image);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageRemoval_PreservesOtherOccurrencesOfTheSameRelationship(bool external) {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("Images");
        var run = paragraph._run!;
        var owner = document.MainDocumentPartRoot;
        string id;
        byte[] expected = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
        if (external) id = owner.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            new Uri("https://example.test/shared.gif")).Id;
        else {
            var imagePart = owner.AddImagePart(ImagePartType.Gif);
            imagePart.FeedData(new MemoryStream(expected));
            id = owner.GetIdOfPart(imagePart);
        }
        run.Append(new W.Picture(new V.Shape(new V.ImageData { RelationshipId = id }) { Id = "First", Style = "width:15pt;height:15pt" }),
            new W.Picture(new V.Shape(new V.ImageData { RelationshipId = id }) { Id = "Second", Style = "width:15pt;height:15pt" }));
        var first = paragraph.Image!;
        Assert.Equal(20, first.Width);
        first.Remove();
        Assert.Single(run.Descendants<V.Shape>());
        var second = paragraph.Image!;
        if (external) Assert.Equal("https://example.test/shared.gif", second.ExternalUri!.AbsoluteUri);
        else Assert.Equal(expected, second.ToBytes());
        second.Remove();
        Assert.Empty(run.Descendants<V.Shape>());
        if (external) Assert.DoesNotContain(owner.ExternalRelationships, relationship => relationship.Id == id);
        else Assert.False(owner.TryGetPartById(id, out _));
    }
}
