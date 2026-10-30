using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

// Synthetic ownership and resource scenarios. The separate native-image test
// exercises independently produced input; this graph is not native writer proof.
internal static class PublisherPictureEditingFixture {
    internal static byte[] Image(OfficeColor color, int width = 2, int height = 2) =>
        OfficeRasterImageEncoder.Encode(new OfficeRasterImage(width, height, color), OfficeImageExportFormat.Png);

    internal static PublisherDocument Create(PublisherReadOptions? options = null, string? restriction = null, bool withWrapping = false) {
        var context = new PublisherParseContext((options ?? new PublisherReadOptions()).Clone(), default);
        var source = new PublisherContents(300, 300, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(20, "A", null, null, true, 3));
        source.Pages.Add(new PublisherSourcePage(10, "One", 20, null));
        source.Pages.Add(new PublisherSourcePage(11, "Two", 20, null));
        var artwork = new PublisherEscherData();
        byte[] image = Image(OfficeColor.FromRgb(255, 0, 0));
        artwork.Images.Add(1, new PublisherImage(1, "image/png", image));
        context.AccountImage(image.Length); context.AccountImageStoreEntry();
        if (restriction == "unresolved-sibling")
            for (int entry = 1; entry < 4; entry++) context.AccountImageStoreEntry();
        foreach (var item in new[] { (Id: 101U, Page: 10U, X: 24D), (Id: 102U, Page: 10U, X: 120D), (Id: 103U, Page: 20U, X: 24D) }) {
            var properties = new List<OfficeArtProperty> { new(0, 0x4104, 1), new(1, 0x0102, 16384) };
            uint flags = 0;
            if (restriction == "unresolved-sibling" && item.Id == 102) properties[0] = new OfficeArtProperty(0, 0x4104, 2);
            if (item.Id == 101) {
                if (restriction == "effects") properties.Add(new OfficeArtProperty(2, 0x013F, 0x00040004));
                if (restriction == "rotation") properties.Add(new OfficeArtProperty(2, 4, 30U * 65536));
                if (restriction == "hidden") properties.Add(new OfficeArtProperty(2, 0x03BF, 0x40004000));
                if (restriction == "reflected") flags = 0x40;
                if (restriction == "unresolved") properties[0] = new OfficeArtProperty(0, 0x4104, 999);
                if (restriction == "invalid-crop") properties.Add(new OfficeArtProperty(2, 0x0103, 65536));
            }
            var shape = new PublisherEscherShape(item.Id, restriction == "nonrectangular" && item.Id == 101 ? 3 : 75, flags, properties) {
                Bounds = restriction == "no-anchor" && item.Id == 101 ? null : new PublisherNativeRectangle(
                    (item.X - 150) * 12700, -120 * 12700, (item.X + 72 - 150) * 12700, -48 * 12700),
                HasPublicationGroupParent = restriction == "grouped" && item.Id == 101
            };
            artwork.Shapes.Add(shape);
            source.Shapes.Add(item.Id, new PublisherSourceShape(new PublisherContentChunk(item.Id, 0x20, 0, item.Page), Array.Empty<PublisherBlock>()));
            source.ObjectPages.Add(item.Id, item.Page);
        }
        var text = new PublisherQuillText();
        if (withWrapping) {
            string body = string.Concat(Enumerable.Repeat("A picture moves and the complete story remains available. ", 12));
            text.Stories.Add(1, new PublisherTextStory(1, body, new[] {
                new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(body, 10, OfficeColor.Black, fontFamily: "Arial") })
            }));
            source.Shapes.Add(201, new PublisherSourceShape(new PublisherContentChunk(201, 1, 0, 10), new[] {
                new PublisherBlock(0x27, 0x20, 0, 4, 1)
            }) { WrapObjects = new uint[] { 101 } });
            source.ObjectPages.Add(201, 10);
            artwork.Shapes.Add(new PublisherEscherShape(201, 202, 0, Array.Empty<OfficeArtProperty>()) {
                Bounds = new PublisherNativeRectangle(-138 * 12700, -138 * 12700, 112 * 12700, 112 * 12700)
            });
        }
        byte[] original = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Image/native-image.pub"));
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out string? error), error);
        return new PublisherModelState(source, text, artwork, context,
            new PublisherNativeSource(original, compound!)).Project(default);
    }
}
