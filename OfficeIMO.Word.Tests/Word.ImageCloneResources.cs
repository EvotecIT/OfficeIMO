using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public sealed class WordImageCloneResourcesTests {
    private static readonly byte[] Gif = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Fact]
    public void Clone_RemapsVmlTextureFillAlongsideImageDataAcrossStories() {
        using var source = WordDocument.Create();
        using var destination = WordDocument.Create();
        var paragraph = source.AddParagraph("Textured image");
        var imagePart = source.MainDocumentPartRoot.AddImagePart(ImagePartType.Gif);
        imagePart.FeedData(new MemoryStream(Gif));
        var texturePart = source.MainDocumentPartRoot.AddImagePart(ImagePartType.Gif);
        texturePart.FeedData(new MemoryStream(Gif));
        var fill = new V.Fill();
        fill.SetAttribute(new OpenXmlAttribute("r", "id", RelationshipsNamespace, source.MainDocumentPartRoot.GetIdOfPart(texturePart)));
        paragraph._run!.Append(new W.Picture(new V.Shape(fill,
            new V.ImageData { RelationshipId = source.MainDocumentPartRoot.GetIdOfPart(imagePart) }) {
                Id = "textured", Style = "width:15pt;height:15pt" }));
        var header = destination.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default);
        header.AddParagraph().AddImage(new MemoryStream(Gif), "existing.gif", 20, 20);
        var target = header.AddParagraph();
        paragraph.Image!.Clone(target);
        var owner = destination.MainDocumentPartRoot.HeaderParts.Single();
        var shape = Assert.Single(target._paragraph.Descendants<V.Shape>());
        var imageId = shape.Descendants<V.ImageData>().Single().RelationshipId!.Value;
        var fillId = shape.Descendants<V.Fill>().Single().GetAttribute("id", RelationshipsNamespace).Value;
        Assert.NotEqual(imageId, fillId);
        Assert.IsType<ImagePart>(owner.GetPartById(imageId));
        Assert.IsType<ImagePart>(owner.GetPartById(fillId));
        using var bytes = destination.ToStream();
        using var reopened = WordDocument.Load(bytes);
        var persistedOwner = reopened.MainDocumentPartRoot.HeaderParts.Single();
        var persistedShape = persistedOwner.Header!.Descendants<V.Shape>().Single();
        Assert.IsType<ImagePart>(persistedOwner.GetPartById(
            persistedShape.Descendants<V.Fill>().Single().GetAttribute("id", RelationshipsNamespace).Value));
    }

    [Fact]
    public void Clone_RemapsDrawingClickAndHoverLinksAcrossStories() {
        using var source = WordDocument.Create();
        using var destination = WordDocument.Create();
        var image = source.AddParagraph().AddImage(new MemoryStream(Gif), "source.gif", 20, 20).Image!;
        var click = source.MainDocumentPartRoot.AddHyperlinkRelationship(new Uri("https://example.test/click"), true);
        var hover = source.MainDocumentPartRoot.AddHyperlinkRelationship(new Uri("https://example.test/hover"), true);
        image._Image.Descendants<PIC.NonVisualDrawingProperties>().Single().Append(
            new A.HyperlinkOnClick { Id = click.Id }, new A.HyperlinkOnHover { Id = hover.Id });

        destination.AddParagraph().AddImage(new MemoryStream(Gif), "existing.gif", 20, 20);
        var target = destination.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph();
        image.Clone(target);
        var owner = destination.MainDocumentPartRoot.HeaderParts.Single();
        var copiedClick = Assert.Single(target._paragraph.Descendants<A.HyperlinkOnClick>());
        var copiedHover = Assert.Single(target._paragraph.Descendants<A.HyperlinkOnHover>());
        Assert.Equal("https://example.test/click", owner.HyperlinkRelationships.Single(item => item.Id == copiedClick.Id!.Value).Uri.ToString());
        Assert.Equal("https://example.test/hover", owner.HyperlinkRelationships.Single(item => item.Id == copiedHover.Id!.Value).Uri.ToString());
        using var package = destination.ToStream();
        using var reopened = WordDocument.Load(package);
        var persisted = reopened.MainDocumentPartRoot.HeaderParts.Single();
        Assert.Equal("https://example.test/click", persisted.HyperlinkRelationships.Single(item =>
            item.Id == persisted.Header!.Descendants<A.HyperlinkOnClick>().Single().Id!.Value).Uri.ToString());
        Assert.Equal("https://example.test/hover", persisted.HyperlinkRelationships.Single(item =>
            item.Id == persisted.Header!.Descendants<A.HyperlinkOnHover>().Single().Id!.Value).Uri.ToString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Clone_AllocatesDrawingIdentifiersAcrossStoriesAndPackages(bool otherDocument) {
        using var source = WordDocument.Create();
        using var destination = WordDocument.Create();
        var image = source.AddParagraph().AddImage(new MemoryStream(Gif), "source.gif", 20, 20).Image!;
        var targetDocument = otherDocument ? destination : source;
        targetDocument.AddParagraph().AddImage(new MemoryStream(Gif), "existing.gif", 20, 20);
        var target = targetDocument.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph();
        image.Clone(target);
        var roots = new OpenXmlElement[] { targetDocument.MainDocumentPartRoot.Document!,
            targetDocument.MainDocumentPartRoot.HeaderParts.Single().Header! };
        var ids = roots.SelectMany(root => root.Descendants<DW.DocProperties>().Select(item => item.Id!.Value)
            .Concat(root.Descendants<PIC.NonVisualDrawingProperties>().Select(item => item.Id!.Value))).ToArray();
        Assert.Equal(ids.Length, ids.Distinct().Count());
    }

    [Fact]
    public void Clone_PreservesReferencedVmlShapeDefinitionDespiteDestinationCollision() {
        using var source = WordDocument.Create();
        using var destination = WordDocument.Create();
        var paragraph = source.AddParagraph("Image");
        var part = source.MainDocumentPartRoot.AddImagePart(ImagePartType.Gif);
        part.FeedData(new MemoryStream(Gif));
        paragraph._run!.Append(new W.Picture(
            new V.Shapetype { Id = "shared-type", CoordinateSize = "21600,21600", EdgePath = "m0,0l21600,21600e" },
            new V.Shape(new V.ImageData { RelationshipId = source.MainDocumentPartRoot.GetIdOfPart(part) }) {
                Id = "source", Type = "#shared-type", Style = "width:15pt;height:15pt" }));
        var header = destination.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default);
        header.AddParagraph("Existing")._run!.Append(new W.Picture(new V.Shapetype { Id = "shared-type", EdgePath = "different" }));
        var target = header.AddParagraph();
        paragraph.Image!.Clone(target);
        var shape = Assert.Single(target._paragraph.Descendants<V.Shape>());
        var definition = Assert.Single(target._paragraph.Descendants<V.Shapetype>());
        Assert.Equal("#" + definition.Id!.Value, shape.Type!.Value);
        Assert.NotEqual("shared-type", definition.Id!.Value);
        Assert.Equal("m0,0l21600,21600e", definition.EdgePath!.Value);
        using var bytes = destination.ToStream();
        using var reopened = WordDocument.Load(bytes);
        Assert.Single(reopened.MainDocumentPartRoot.HeaderParts.Single().Header.Descendants<V.Shapetype>(),
            item => item.EdgePath?.Value == "m0,0l21600,21600e");
    }

    [Theory]
    [InlineData("http://schemas.microsoft.com/office/drawing/2016/SVG/main")]
    [InlineData("http://schemas.microsoft.com/office/drawing/2010/main")]
    public void Clone_RemapsSvgExtensionToDestinationImagePart(string svgNamespace) {
        using var source = WordDocument.Create();
        using var destination = WordDocument.Create();
        var image = source.AddParagraph().AddImage(new MemoryStream(Gif), "source.gif", 20, 20).Image!;
        var svgBytes = System.Text.Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'><rect width='20' height='20'/></svg>");
        var svgPart = source.MainDocumentPartRoot.AddImagePart(ImagePartType.Svg);
        svgPart.FeedData(new MemoryStream(svgBytes));
        const string relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var extension = new OpenXmlUnknownElement("svg", "svgBlip", svgNamespace);
        extension.SetAttribute(new OpenXmlAttribute("r", "embed", relationships, source.MainDocumentPartRoot.GetIdOfPart(svgPart)));
        image._Image.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Single().Append(
            new DocumentFormat.OpenXml.Drawing.BlipExtensionList(new DocumentFormat.OpenXml.Drawing.BlipExtension(extension) { Uri = "{96DAC541-7B7A-43D3-8B79-8F7C33B92B69}" }));
        var header = destination.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default);
        header.AddParagraph().AddImage(new MemoryStream(Gif), "existing.gif", 20, 20);
        var target = header.AddParagraph();
        image.Clone(target);
        var owner = destination.MainDocumentPartRoot.HeaderParts.Single();
        var copied = target._paragraph.Descendants().Single(item => item.LocalName == "svgBlip");
        var id = copied.GetAttribute("embed", relationships).Value;
        var copiedPart = Assert.IsType<ImagePart>(owner.GetPartById(id));
        using var bytes = copiedPart.GetStream();
        using var output = new MemoryStream();
        bytes.CopyTo(output);
        Assert.Equal(svgBytes, output.ToArray());
        using var package = destination.ToStream();
        using var reopened = WordDocument.Load(package);
        var persistedOwner = reopened.MainDocumentPartRoot.HeaderParts.Single();
        var persistedSvg = persistedOwner.Header!.Descendants().Single(item => item.LocalName == "svgBlip");
        Assert.Equal("image/svg+xml", Assert.IsType<ImagePart>(
            persistedOwner.GetPartById(persistedSvg.GetAttribute("embed", relationships).Value)).ContentType);
    }
}
