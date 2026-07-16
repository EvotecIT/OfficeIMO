using System.Text;

namespace OfficeIMO.OneNote.Tests;

public sealed class ImageCompatibilityTests {
    [Fact]
    public void ReadsWebPictureFallbackAndPreservesBothNativeReferences() {
        byte[] webPayload = { 9, 8, 7, 6 };
        OneNoteWriteGraph graph = CreateImageGraph();
        OneNoteWriteObjectSpace pageSpace = graph.ObjectSpaces[1];
        OneNoteWriteObject imageObject = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidImageNode);
        OneNoteExtendedGuid primaryId = Assert.Single(Property(imageObject, OneNoteSchema.PictureContainer).References);
        MakePayloadUnresolved(pageSpace, primaryId);

        OneNoteExtendedGuid webId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 4);
        Guid webDataId = Guid.NewGuid();
        pageSpace.Objects.Add(new OneNoteWriteObject(
            webId,
            OneNoteSchema.JcidPictureData,
            FileDataProperties(webDataId, ".png"),
            webPayload,
            webDataId,
            ".png"));
        ReplaceObject(pageSpace, imageObject, imageObject.Properties.Concat(new[] {
            new OneNoteWriteProperty(OneNoteSchema.WebPictureContainer14, references: new[] { webId })
        }));

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)));
        OneNoteImage image = GetImage(loaded);

        Assert.Equal(webPayload, image.Payload!.ToArray(16));
        Assert.Equal(primaryId, image.PictureContainerObjectId);
        Assert.Equal(webId, image.WebPictureContainerObjectId);
        Assert.True(image.PayloadUsesWebPictureContainer);

        byte[] rewritten = OneNoteSectionWriter.Write(loaded);
        AssertPreservedReferences(rewritten, primaryId, webId, webPayload);
        Assert.Equal(webPayload, GetImage(OneNoteSectionReader.Read(new MemoryStream(rewritten))).Payload!.ToArray(16));

        byte[] canonical = OneNoteSectionWriter.Write(loaded, new OneNoteWriterOptions { PreserveUnknownData = false });
        OneNoteImage canonicalImage = GetImage(OneNoteSectionReader.Read(new MemoryStream(canonical)));
        Assert.Equal(webPayload, canonicalImage.Payload!.ToArray(16));
        Assert.False(canonicalImage.PayloadUsesWebPictureContainer);
    }

    [Fact]
    public void PreservesUnresolvedPictureContainerWithoutInventingPayload() {
        OneNoteWriteGraph graph = CreateImageGraph();
        OneNoteWriteObjectSpace pageSpace = graph.ObjectSpaces[1];
        OneNoteWriteObject imageObject = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidImageNode);
        OneNoteExtendedGuid primaryId = Assert.Single(Property(imageObject, OneNoteSchema.PictureContainer).References);
        MakePayloadUnresolved(pageSpace, primaryId);

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)));
        OneNoteImage image = GetImage(loaded);

        Assert.Null(image.Payload);
        Assert.Equal(primaryId, image.PictureContainerObjectId);

        byte[] rewritten = OneNoteSectionWriter.Write(loaded);
        OneNoteSection roundTrip = OneNoteSectionReader.Read(new MemoryStream(rewritten));
        Assert.Null(GetImage(roundTrip).Payload);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(loaded, new OneNoteWriterOptions { PreserveUnknownData = false }));
        Assert.Equal("ONENOTE_WRITE_MISSING_PAYLOAD", exception.Code);
    }

    [Fact]
    public void PreservesPictureReferenceWhenTheTargetObjectIsMissing() {
        OneNoteWriteGraph graph = CreateImageGraph();
        OneNoteWriteObjectSpace pageSpace = graph.ObjectSpaces[1];
        OneNoteWriteObject imageObject = Assert.Single(pageSpace.Objects, item => item.Jcid == OneNoteSchema.JcidImageNode);
        OneNoteExtendedGuid primaryId = Assert.Single(Property(imageObject, OneNoteSchema.PictureContainer).References);
        OneNoteWriteObject payload = Assert.Single(pageSpace.Objects, item => item.Id.Equals(primaryId));
        Assert.True(pageSpace.Objects.Remove(payload));

        OneNoteSection loaded = OneNoteSectionReader.Read(new MemoryStream(OneNoteRevisionStoreWriter.Write(graph)));
        OneNoteImage image = GetImage(loaded);

        Assert.Null(image.Payload);
        Assert.Equal(primaryId, image.PictureContainerObjectId);

        byte[] rewritten = OneNoteSectionWriter.Write(loaded);
        Assert.Equal(primaryId, PictureReference(rewritten));

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(loaded, new OneNoteWriterOptions { PreserveUnknownData = false }));
        Assert.Equal("ONENOTE_WRITE_MISSING_PAYLOAD", exception.Code);
    }

    private static OneNoteWriteGraph CreateImageGraph() {
        var section = new OneNoteSection { Name = "Images" };
        var page = new OneNotePage { Title = "Web fallback" };
        page.DirectContent.Add(new OneNoteImage {
            FileName = "preview.png",
            WidthHalfInches = 21.3333339691162,
            HeightHalfInches = 10.6666669845581,
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3, 4 })
        });
        section.Pages.Add(page);
        return new OneNoteWriteGraphBuilder().BuildSection(section);
    }

    private static void MakePayloadUnresolved(OneNoteWriteObjectSpace pageSpace, OneNoteExtendedGuid payloadId) {
        OneNoteWriteObject payload = Assert.Single(pageSpace.Objects, item => item.Id.Equals(payloadId));
        int index = pageSpace.Objects.IndexOf(payload);
        pageSpace.Objects[index] = new OneNoteWriteObject(
            payload.Id,
            payload.Jcid,
            payload.Properties,
            blob: null,
            fileDataId: null,
            fileExtension: payload.FileExtension);
    }

    private static void AssertPreservedReferences(
        byte[] data,
        OneNoteExtendedGuid primaryId,
        OneNoteExtendedGuid webId,
        byte[] expectedWebPayload) {
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));
        var materializer = new OneNoteObjectSpaceMaterializer(store);
        OneNoteMaterializedObjectSpace pageSpace = materializer.FindCurrentSpaceByRootJcid(
            OneNoteSchema.JcidPageManifestNode,
            "TEST_PAGE",
            "Expected a page space.");
        OneNoteRevisionStoreObject image = Assert.Single(pageSpace.Objects, item => item.Jcid.Value == OneNoteSchema.JcidImageNode);
        Assert.Equal(primaryId, Assert.Single(OneNoteSemanticMapper.GetReferences(image, OneNoteSchema.PictureContainer)));
        Assert.Equal(webId, Assert.Single(OneNoteSemanticMapper.GetReferences(image, OneNoteSchema.WebPictureContainer14)));

        OneNoteRevisionStoreObject primary = Assert.IsType<OneNoteRevisionStoreObject>(pageSpace.GetObject(primaryId));
        Assert.False(materializer.TryResolveFileData(primary, out _, out _));
        OneNoteRevisionStoreObject web = Assert.IsType<OneNoteRevisionStoreObject>(pageSpace.GetObject(webId));
        Assert.True(materializer.TryResolveFileData(web, out _, out OneNoteBinaryPayload? payload));
        Assert.Equal(expectedWebPayload, payload!.ToArray(16));
    }

    private static OneNoteExtendedGuid PictureReference(byte[] data) {
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));
        var materializer = new OneNoteObjectSpaceMaterializer(store);
        OneNoteMaterializedObjectSpace pageSpace = materializer.FindCurrentSpaceByRootJcid(
            OneNoteSchema.JcidPageManifestNode,
            "TEST_PAGE",
            "Expected a page space.");
        OneNoteRevisionStoreObject image = Assert.Single(pageSpace.Objects, item => item.Jcid.Value == OneNoteSchema.JcidImageNode);
        return Assert.Single(OneNoteSemanticMapper.GetReferences(image, OneNoteSchema.PictureContainer));
    }

    private static OneNoteImage GetImage(OneNoteSection section) =>
        Assert.IsType<OneNoteImage>(Assert.Single(Assert.Single(Assert.Single(section.Pages).Outlines).Children));

    private static IReadOnlyList<OneNoteWriteProperty> FileDataProperties(Guid dataId, string extension) =>
        new[] {
            new OneNoteWriteProperty(OneNoteSchema.FileDataReference, data: dataId.ToByteArray()),
            new OneNoteWriteProperty(OneNoteSchema.FileDataExtension, data: Encoding.Unicode.GetBytes(extension + "\0"))
        };

    private static OneNoteWriteProperty Property(OneNoteWriteObject item, uint id) =>
        Assert.Single(item.Properties, property => (property.RawId & 0x7FFFFFFFU) == id);

    private static void ReplaceObject(
        OneNoteWriteObjectSpace space,
        OneNoteWriteObject source,
        IEnumerable<OneNoteWriteProperty> properties,
        byte[]? blob = null,
        Guid? fileDataId = null,
        string? extension = null) {
        int index = space.Objects.IndexOf(source);
        space.Objects[index] = new OneNoteWriteObject(
            source.Id,
            source.Jcid,
            properties,
            blob ?? source.Blob,
            fileDataId ?? source.FileDataId,
            extension ?? source.FileExtension);
    }
}
