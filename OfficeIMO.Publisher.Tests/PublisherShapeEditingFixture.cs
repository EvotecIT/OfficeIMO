using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

// Synthetic ownership, wrapping and admission variants; native Rectangle and
// native-authored controls provide the independent producer input separately.
internal static class PublisherShapeEditingFixture {
    internal static PublisherDocument Create(string? restriction = null, bool withWrapping = false,
        PublisherReadOptions? options = null, int hiddenObjects = 0, bool exhaustedIds = false) {
        var context = new PublisherParseContext((options ?? new PublisherReadOptions()).Clone(), default);
        var source = new PublisherContents(300, 300, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(20, "A", null, null, true, restriction == "unqualified-master" ? 2U : 3U));
        source.Pages.Add(new PublisherSourcePage(10, "One", 20, null));
        source.Pages.Add(new PublisherSourcePage(11, "Two", 20, null));
        var artwork = new PublisherEscherData();
        foreach (var item in new[] { (Id: 101U, Page: 10U, X: 24D), (Id: 102U, Page: 10U, X: 120D), (Id: 103U, Page: 20U, X: 24D) }) {
            bool plainLine = item.Id == 101 && restriction == "plain-line";
            var properties = new List<OfficeArtProperty> {
                new(0, 0x0181, 0x000000FF), new(1, 0x01BF, 0x00100010),
                new(2, 0x01C0, 0), new(3, 0x01FF, 0x00080008), new(4, 0x01CB, 12700)
            };
            if (item.Id == 101) {
                if (restriction == "gradient") properties.Add(new OfficeArtProperty(5, 0x0180, 4));
                if (restriction == "shadow") properties.Add(new OfficeArtProperty(5, 0x023F, 0x00020002));
                if (restriction == "perspective") properties.Add(new OfficeArtProperty(5, 0x0285, 65536));
                if (restriction == "rotation") properties.Add(new OfficeArtProperty(5, 4, 30 * 65536));
                if (restriction == "hidden") properties.Add(new OfficeArtProperty(5, 0x03BF, 0x40004000));
                if (restriction == "unresolved-color") properties[0] = new OfficeArtProperty(0, 0x0181, 0x080000FF);
                if (restriction == "dash") properties.Add(new OfficeArtProperty(5, 0x01CE, 1));
                if (restriction == "custom-path") properties.Add(new OfficeArtProperty(5, 0x0145, 1));
            }
            double top = item.Page == 20 ? 160 : 30;
            var shape = new PublisherEscherShape(item.Id, plainLine ? 20 : restriction == "unknown-preset" && item.Id == 101 ? 100 : 1,
                restriction == "reflected" && item.Id == 101 ? 0x40U : 0U, properties) {
                Bounds = restriction == "no-anchor" && item.Id == 101 ? null : new PublisherNativeRectangle(
                    (item.X - 150) * 12700, (top - 150) * 12700, (item.X + 72 - 150) * 12700, (top + (plainLine ? 0 : 72) - 150) * 12700),
                HasPublicationGroupParent = restriction == "grouped" && item.Id == 101
            };
            source.Shapes.Add(item.Id, new PublisherSourceShape(new PublisherContentChunk(item.Id, plainLine ? 0x20U : 1U, 0, item.Page), Array.Empty<PublisherBlock>()));
            source.ObjectPages.Add(item.Id, item.Page); artwork.Shapes.Add(shape);
        }
        var text = new PublisherQuillText();
        if (withWrapping) {
            string body = string.Concat(Enumerable.Repeat("A vector shape moves while the complete story remains available. ", 12));
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
        for (int index = 0; index < hiddenObjects || index == 0 && exhaustedIds; index++) {
            uint id = exhaustedIds ? uint.MaxValue : 1000U + (uint)index;
            source.Shapes.Add(id, new PublisherSourceShape(new PublisherContentChunk(id, 1, 0, 10), Array.Empty<PublisherBlock>()));
            source.ObjectPages.Add(id, 10);
            artwork.Shapes.Add(new PublisherEscherShape(id, 1, 0, new[] { new OfficeArtProperty(0, 0x03BF, 0x40004000) }));
        }
        byte[] bytes = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Rectangle/native-rectangle.pub"));
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        return new PublisherModelState(source, text, artwork, context, new PublisherNativeSource(bytes, compound!)).Project(default);
    }
}
