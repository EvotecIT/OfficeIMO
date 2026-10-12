using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;
public sealed class PublisherCustomStrokePlacementTests {
    [Theory]
    [InlineData(true, 0)]
    [InlineData(false, 0)]
    [InlineData(true, 20)]
    [InlineData(false, 20)]
    public void One_dimensional_native_freeforms_keep_their_custom_stroke(bool horizontal, int preset) {
        var source = new PublisherContents(100, 100, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        source.Shapes.Add(2, new PublisherSourceShape(new PublisherContentChunk(2, 1, 0, 1), Array.Empty<PublisherBlock>()));
        source.ObjectPages.Add(2, 1);
        byte[] vertices = PublisherCustomPathTests.Vertices(new[] { (0, 0), (21600, 21600) });
        byte[] segments = PublisherCustomPathTests.Segments(new ushort[] { 0x4000, 0x0001, 0x8000 });
        var escher = new PublisherEscherData();
        escher.Shapes.Add(new PublisherEscherShape(2, preset, 0, new[] {
            new OfficeArtProperty(0, 0x8145, (uint)vertices.Length, vertices.Length, complexData: vertices),
            new OfficeArtProperty(1, 0x8146, (uint)segments.Length, segments.Length, complexData: segments),
            new OfficeArtProperty(2, 0x01C0, 0xFF)
        }) { Bounds = new PublisherNativeRectangle(0, 0, horizontal ? 127000 : 0, horizontal ? 0 : 127000) });
        PublisherDocument document = new PublisherPageProjector(source, new PublisherQuillText(), escher,
            new PublisherParseContext(new PublisherReadOptions(), default)).Project();
        OfficeShape path = Assert.Single(PublisherNativeTests.Elements(document.Pages[0].Drawing).OfType<OfficeDrawingShape>(),
            item => item.Shape.Kind == OfficeShapeKind.Path).Shape;
        Assert.Equal(new OfficePoint(horizontal ? 10 : 0, horizontal ? 0 : 10), path.PathCommands[1].Point);
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_INVALID" || item.Code == "PUB_CUSTOM_PATH_GEOMETRY_UNSUPPORTED");
    }
}
