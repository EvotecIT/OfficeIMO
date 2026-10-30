using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

// Bounded model/resource variants built from decoded independent input. They
// qualify admission and identity behavior, not native writing or native editing.
internal static class PublisherFrameCreationFixture {
    internal static (PublisherDocument Document, uint ReservedId) Load(bool danglingReference = false, int extraHiddenObjects = 0,
        PublisherReadOptions? options = null, bool mixedPicture = false) {
        byte[] bytes = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TextBlank/native-textblank.pub"));
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        var context = new PublisherParseContext((options ?? new PublisherReadOptions()).Clone(), default);
        var data = new PublisherBinaryData(compound!.Streams["Contents"], "Contents");
        PublisherContents source = new PublisherContentsReader(data, context).Read();
        PublisherTextFrameCatalogReader.Read(data, source, context);
        PublisherQuillText text = new PublisherQuillTextReader(new PublisherBinaryData(compound.Streams["Quill/QuillSub/CONTENTS"], "Quill"), context).Read(source.Palette);
        compound.Streams.TryGetValue("Escher/EscherDelayStm", out byte[]? delay);
        PublisherEscherData artwork = new PublisherEscherReader(new PublisherBinaryData(compound.Streams["Escher/EscherStm"], "Escher"),
            delay == null ? null : new PublisherBinaryData(delay, "EscherDelay"), context).Read();
        PublisherSourceShape original = source.Shapes.Values.Single(shape => shape.Value(0x27).HasValue);
        uint pageId = source.ObjectPages[original.Chunk.Id];
        uint reserved = source.Chunks.Keys.Concat(source.Shapes.Keys).DefaultIfEmpty(0U).Max() + 1;
        if (danglingReference) original.WrapObjects = new[] { reserved };
        if (mixedPicture) {
            PublisherEscherShape existing = artwork.Shapes.Single(shape => shape.Id == original.Chunk.Id);
            var properties = existing.Properties.Concat(new[] { new OfficeArtProperty(existing.Properties.Count, 0x4104, 1) }).ToArray();
            var replacement = new PublisherEscherShape(existing.Id, 1, existing.Flags, properties, existing.HasParentDrawingSpace) {
                Bounds = existing.Bounds, GroupCoordinates = existing.GroupCoordinates, GroupTransform = existing.GroupTransform,
                LayoutTransform = existing.LayoutTransform, HasPublicationGroupParent = existing.HasPublicationGroupParent, HiddenByGroup = existing.HiddenByGroup
            };
            artwork.Shapes[artwork.Shapes.IndexOf(existing)] = replacement;
            byte[] image = PublisherPictureEditingFixture.Image(OfficeColor.FromRgb(0, 160, 40));
            artwork.Images.Add(1, new PublisherImage(1, "image/png", image)); context.AccountImage(image.Length); context.AccountImageStoreEntry();
        }
        for (int index = 0; index < extraHiddenObjects; index++) {
            uint id = reserved + (uint)index;
            source.Shapes.Add(id, new PublisherSourceShape(new PublisherContentChunk(id, 1, 0, pageId), Array.Empty<PublisherBlock>()));
            source.ObjectPages.Add(id, pageId);
            artwork.Shapes.Add(new PublisherEscherShape(id, 1, 0, new[] { new OfficeArtProperty(0, 0x03BF, 0x40004000) }));
        }
        return (new PublisherModelState(source, text, artwork, context, new PublisherNativeSource(bytes, compound)).Project(default), reserved);
    }
}
