using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Tests;

public sealed class WordImageVmlOwnershipTests {
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
