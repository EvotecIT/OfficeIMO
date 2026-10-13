using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherGroupCanvasTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Empty_grouped_children_do_not_consume_effect_canvases(bool textFrame) {
        var source = new PublisherContents(612, 792, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        var text = new PublisherQuillText();
        var escher = new PublisherEscherData();
        for (uint id = 2; id < 106; id++) {
            text.Stories.Add(id + 500, new PublisherTextStory(id + 500, "", Array.Empty<OfficeRichTextParagraph>()));
            source.Shapes.Add(id, new PublisherSourceShape(new PublisherContentChunk(id, 1, 0, 1),
                textFrame ? new[] { new PublisherBlock(0x27, 0, 0, 0, id + 500) } : Array.Empty<PublisherBlock>()));
            source.ObjectPages.Add(id, 1);
            escher.Shapes.Add(new PublisherEscherShape(id, 1, 0, new[] {
                new OfficeArtProperty(0, 0x01FF, 0x00080000), new OfficeArtProperty(1, 0x01BF, 0x00100000)
            }) {
                Bounds = new PublisherNativeRectangle(-306 * 12700, -396 * 12700, 306 * 12700, 396 * 12700),
                GroupTransform = OfficeTransform.Translate(12700, 12700)
            });
        }
        PublisherDocument document = new PublisherPageProjector(source, text, escher,
            new PublisherParseContext(new PublisherReadOptions(), default)).Project();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(document.Pages[0].Drawing);
        Assert.Equal(OfficeColor.White, raster.GetPixel(300, 300));
        Assert.DoesNotContain(PublisherNativeTests.Elements(document.Pages[0].Drawing), item => item is OfficeDrawingEffectGroup);
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_OBJECT_NOT_PROJECTED");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Grouped_text_retains_glyph_and_decoration_ink_outside_its_frame(bool condensed) {
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(TextPaint(false, condensed).Pages[0].Drawing);
        OfficeRasterImage grouped = OfficeDrawingRasterRenderer.Render(TextPaint(true, condensed).Pages[0].Drawing);
        int painted = 0, missing = 0;
        for (int y = 60; y < 200; y++)
            for (int x = 60; x < 220; x++) {
                OfficeColor expected = reference.GetPixel(x, y);
                if (expected.R == 255) continue;
                painted++;
                if (grouped.GetPixel(x, y).R > expected.R) missing++;
            }
        Assert.True(painted > 0);
        Assert.Equal(0, missing);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Collapsed_group_text_frames_preserve_the_story_without_projection(bool horizontal) {
        PublisherDocument document = TextPaint(true, false, horizontal);
        Assert.Empty(document.Pages[0].TextFrames);
        Assert.Equal("j", Assert.Single(document.TextStories).Text);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_CHAIN_FRAME_UNAVAILABLE");
    }

    [Fact]
    public void Collapsed_wrap_obstacles_do_not_abort_available_text_frames() {
        PublisherDocument document = TextPaint(true, false, collapsedObstacle: true);
        Assert.Single(document.Pages[0].TextFrames);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_TEXT_WRAP_REFERENCE_UNRESOLVED");
    }

    [Fact]
    public void Collapsed_group_lines_do_not_paint_minimum_sized_artwork() {
        var source = new PublisherContents(612, 792, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        source.Shapes.Add(2, new PublisherSourceShape(new PublisherContentChunk(2, 1, 0, 1), Array.Empty<PublisherBlock>()));
        source.ObjectPages.Add(2, 1);
        var escher = new PublisherEscherData();
        escher.Shapes.Add(new PublisherEscherShape(2, 20, 0, new[] {
            new OfficeArtProperty(0, 0x01C0, 0x000000FF), new OfficeArtProperty(1, 0x01CB, 12700)
        }) { Bounds = new PublisherNativeRectangle(0, 0, 12700, 12700),
            GroupTransform = OfficeTransform.Scale(0, 1), LayoutTransform = OfficeTransform.Scale(0, 1) });
        PublisherDocument document = new PublisherPageProjector(source, new PublisherQuillText(), escher,
            new PublisherParseContext(new PublisherReadOptions(), default)).Project();
        Assert.DoesNotContain(PublisherNativeTests.Elements(document.Pages[0].Drawing), item => item is OfficeDrawingShape drawing && drawing.Shape.Kind == OfficeShapeKind.Line);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_DEGENERATE_OBJECT_OMITTED");
    }

    private static PublisherDocument TextPaint(bool grouped, bool condensed, bool? collapseHorizontal = null, bool collapsedObstacle = false) {
        var source = new PublisherContents(612, 792, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        source.Shapes.Add(2, new PublisherSourceShape(new PublisherContentChunk(2, 1, 0, 1),
            new[] { new PublisherBlock(0x27, 0, 0, 0, 3) }));
        source.ObjectPages.Add(2, 1);
        string value = condensed ? "j\nj" : "j";
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(value, 36, OfficeColor.Black,
            italic: true, underline: true, fontFamily: "Arial") }, lineHeight: condensed ? 18 : null);
        var text = new PublisherQuillText();
        text.Stories.Add(3, new PublisherTextStory(3, value, new[] { paragraph }));
        double x = grouped ? 100 : 101;
        var escher = new PublisherEscherData();
        escher.Shapes.Add(new PublisherEscherShape(2, 1, 0, new[] {
            new OfficeArtProperty(0, 0x81, 0), new OfficeArtProperty(1, 0x82, 0),
            new OfficeArtProperty(2, 0x83, 0), new OfficeArtProperty(3, 0x84, 0),
            new OfficeArtProperty(4, 0x01FF, 0x00080000), new OfficeArtProperty(5, 0x01BF, 0x00100000)
        }) {
            Bounds = new PublisherNativeRectangle((x - 306) * 12700, (x - 396) * 12700,
                (x + 25 - 306) * 12700, (x + 72 - 396) * 12700),
            GroupTransform = collapseHorizontal.HasValue ? OfficeTransform.Scale(collapseHorizontal.Value ? 0 : 1, collapseHorizontal.Value ? 1 : 0) : grouped ? OfficeTransform.Translate(12700, 12700) : OfficeTransform.Identity,
            LayoutTransform = collapseHorizontal.HasValue ? OfficeTransform.Scale(collapseHorizontal.Value ? 0 : 1, collapseHorizontal.Value ? 1 : 0) : OfficeTransform.Identity
        });
        if (collapsedObstacle) {
            source.Shapes[2].WrapObjects = new[] { 4U };
            source.Shapes.Add(4, new PublisherSourceShape(new PublisherContentChunk(4, 1, 0, 1), Array.Empty<PublisherBlock>()));
            source.ObjectPages.Add(4, 1);
            escher.Shapes.Add(new PublisherEscherShape(4, 1, 0, Array.Empty<OfficeArtProperty>()) {
                Bounds = new PublisherNativeRectangle(0, 0, 12700, 12700),
                GroupTransform = OfficeTransform.Scale(0, 1), LayoutTransform = OfficeTransform.Scale(0, 1)
            });
        }
        return new PublisherPageProjector(source, text, escher,
            new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None)).Project();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Grouped_shapes_retain_wide_strokes_and_blurred_shadows(bool shadow) {
        OfficeRasterImage reference = OfficeDrawingRasterRenderer.Render(WidePaint(false, shadow).Pages[0].Drawing);
        OfficeRasterImage grouped = OfficeDrawingRasterRenderer.Render(WidePaint(true, shadow).Pages[0].Drawing);
        int painted = 0, missing = 0;
        for (int y = 0; y < 300; y++)
            for (int x = 0; x < 300; x++) {
                OfficeColor expected = reference.GetPixel(x, y);
                if (expected.R == 255) continue;
                painted++;
                if (grouped.GetPixel(x, y).R > expected.R) missing++;
            }
        Assert.True(painted > 0);
        Assert.Equal(0, missing);
    }

    internal static PublisherDocument WidePaint(bool grouped, bool shadow) {
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

    [Theory]
    [InlineData(0)]
    [InlineData(100)]
    public void Many_small_grouped_objects_render_with_the_default_pixel_budget(int disabledStrokeWidth) {
        PublisherDocument document = SmallGroupedObjects(disabledStrokeWidth);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(document.Pages.Single().Drawing);
        for (int index = 0; index < 120; index++) {
            OfficeColor pixel = raster.GetPixel(22 + index % 12 * 40, 22 + index / 12 * 40);
            Assert.Equal(255, pixel.R);
            Assert.Equal(0, pixel.G);
            Assert.Equal(0, pixel.B);
            Assert.Equal(255, pixel.A);
        }
    }

    [Fact]
    public void Transparent_grouped_line_markers_remain_within_the_default_pixel_budget() {
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(SmallGroupedObjects(100, transparentMarkers: true).Pages[0].Drawing);
        int painted = 0;
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                if (raster.GetPixel(x, y) != OfficeColor.White) painted++;
        Assert.Equal(0, painted);
    }

    private static PublisherDocument SmallGroupedObjects(int disabledStrokeWidth = 0, bool transparentMarkers = false) {
        var source = new PublisherContents(612, 792, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        var escher = new PublisherEscherData();
        for (uint index = 0; index < 120; index++) {
            uint id = index + 2;
            double x = 20 + index % 12 * 40, y = 20 + index / 12 * 40;
            source.Shapes.Add(id, new PublisherSourceShape(new PublisherContentChunk(id, 1, 0, 1), Array.Empty<PublisherBlock>()));
            source.ObjectPages.Add(id, 1);
            escher.Shapes.Add(new PublisherEscherShape(id, transparentMarkers ? 20 : 1, 0, new[] {
                new OfficeArtProperty(0, 0x0181, 0x000000FF),
                new OfficeArtProperty(1, 0x01FF, transparentMarkers ? 0x00080008U : 0x00080000U),
                new OfficeArtProperty(2, 0x01CB, (uint)(disabledStrokeWidth * 12700)),
                new OfficeArtProperty(3, 0x01C0, 0),
                new OfficeArtProperty(4, 0x01C1, transparentMarkers ? 0U : 65536U),
                new OfficeArtProperty(5, 0x01D0, transparentMarkers ? 1U : 0U),
                new OfficeArtProperty(6, 0x01D1, transparentMarkers ? 1U : 0U)
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
