using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGroupCanvasTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Grouped_shapes_retain_square_caps_and_blurred_miter_shadows(bool shadow) {
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(WidePaint(false, shadow).Pages[0].Drawing);
        OfficeRasterImage grouped = OfficeDrawingRasterRenderer.Render(WidePaint(true, shadow).Pages[0].Drawing);
        int x = shadow ? 121 : 40, y = shadow ? 61 : 101;
        OfficeColor expected = reference.GetPixel(x, y);
        Assert.True(expected.R < 255);
        Assert.Equal(expected, grouped.GetPixel(x, y));
    }

    private static PublisherDocument WidePaint(bool grouped, bool shadow) {
        var source = new PublisherContents(612, 792, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        source.Shapes.Add(2, new PublisherSourceShape(new PublisherContentChunk(2, 1, 0, 1), Array.Empty<PublisherBlock>()));
        source.ObjectPages.Add(2, 1);
        double x = grouped ? 100 : 101, y = x, width = shadow ? 40 : 100;
        var properties = new List<OfficeArtProperty> {
            new(0, 0x0181, 0x000000FF), new(1, 0x01C0, 0),
            new(2, 0x01CB, (uint)((shadow ? 2 : 100) * 12700)),
            new(3, 0x01CC, (uint)((shadow ? 8 : 1) << 16)),
            new(4, 0x01D6, 1), new(5, 0x01D7, 1)
        };
        if (shadow) properties.AddRange(new[] {
            new OfficeArtProperty(6, 0x023F, 0x00020002), new OfficeArtProperty(7, 0x0201, 0),
            new OfficeArtProperty(8, 0x021C, 127000)
        });
        var escher = new PublisherEscherData();
        escher.Shapes.Add(new PublisherEscherShape(2, shadow ? 5 : 20, 0, properties) {
            Bounds = new PublisherNativeRectangle((x - 306) * 12700, (y - 396) * 12700,
                (x + width - 306) * 12700, (y + 100 - 396) * 12700),
            GroupTransform = grouped ? OfficeTransform.Translate(12700, 12700) : OfficeTransform.Identity
        });
        return new PublisherPageProjector(source, new PublisherQuillText(), escher,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Project();
    }

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
