using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGroupCanvasTests {
    [Fact]
    public void Many_small_grouped_objects_render_with_the_default_pixel_budget() {
        PublisherDocument document = SmallGroupedObjects();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(document.Pages.Single().Drawing);
        for (int index = 0; index < 120; index++) {
            OfficeColor pixel = raster.GetPixel(22 + index % 12 * 40, 22 + index / 12 * 40);
            Assert.Equal(255, pixel.R);
            Assert.Equal(0, pixel.G);
            Assert.Equal(0, pixel.B);
            Assert.Equal(255, pixel.A);
        }
    }

    private static PublisherDocument SmallGroupedObjects() {
        var source = new PublisherContents(612, 792, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        var escher = new PublisherEscherData();
        for (uint index = 0; index < 120; index++) {
            uint id = index + 2;
            double x = 20 + index % 12 * 40, y = 20 + index / 12 * 40;
            source.Shapes.Add(id, new PublisherSourceShape(new PublisherContentChunk(id, 1, 0, 1), Array.Empty<PublisherBlock>()));
            source.ObjectPages.Add(id, 1);
            escher.Shapes.Add(new PublisherEscherShape(id, 1, 0, new[] {
                new OfficeArtProperty(0, 0x0181, 0x000000FF),
                new OfficeArtProperty(1, 0x01FF, 0x00080000)
            }) {
                Bounds = new PublisherNativeRectangle((x - 306) * 12700, (y - 396) * 12700,
                    (x + 4 - 306) * 12700, (y + 4 - 396) * 12700),
                GroupTransform = OfficeTransform.Translate(12700, 12700)
            });
        }
        return new PublisherPageProjector(source, new PublisherQuillText(), escher,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Project();
    }
}
