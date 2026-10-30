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
        OfficeShape path = Assert.Single(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingShape>(),
            item => item.Shape.Kind == OfficeShapeKind.Path).Shape;
        Assert.Equal(new OfficePoint(horizontal ? 10 : 0, horizontal ? 0 : 10), path.PathCommands[1].Point);
        PublisherShape inspected = Assert.Single(document.Pages[0].Shapes);
        Assert.True(inspected.HasPlacement);
        Assert.Equal(50, inspected.X); Assert.Equal(50, inspected.Y);
        Assert.Equal(horizontal ? 10 : 0, inspected.Width);
        Assert.Equal(horizontal ? 0 : 10, inspected.Height);
        Assert.Equal(OfficeTransform.Translate(50, 50), inspected.PageTransform);
        Assert.False(inspected.CanEdit);
        Assert.Contains("geometry", inspected.EditLimitation!);
        Assert.Throws<NotSupportedException>(() => inspected.CreateShape());
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_INVALID" || item.Code == "PUB_CUSTOM_PATH_GEOMETRY_UNSUPPORTED");
    }
    [Theory]
    [InlineData(true, 0)]
    [InlineData(false, 0)]
    [InlineData(true, 2)]
    [InlineData(false, 2)]
    [InlineData(true, 4)]
    [InlineData(false, 4)]
    public void One_dimensional_unsupported_custom_paths_keep_a_bounded_preset_fallback(bool horizontal, int preset) {
        var source = new PublisherContents(100, 100, new Dictionary<uint, PublisherContentChunk>());
        source.Pages.Add(new PublisherSourcePage(1, "", null, null));
        source.Shapes.Add(2, new PublisherSourceShape(new PublisherContentChunk(2, 1, 0, 1), Array.Empty<PublisherBlock>()));
        source.ObjectPages.Add(2, 1);
        byte[] vertices = PublisherCustomPathTests.Vertices(new[] { (0, 0), (21600, 21600) });
        byte[] segments = PublisherCustomPathTests.Segments(new ushort[] { 0x4000, 0xC000, 0x8000 });
        var escher = new PublisherEscherData();
        escher.Shapes.Add(new PublisherEscherShape(2, preset, 0, new[] {
            new OfficeArtProperty(0, 0x8145, (uint)vertices.Length, vertices.Length, complexData: vertices),
            new OfficeArtProperty(1, 0x8146, (uint)segments.Length, segments.Length, complexData: segments),
            new OfficeArtProperty(2, 0x01C0, 0xFF),
            new OfficeArtProperty(3, 0x0180, 4),
            new OfficeArtProperty(4, 0x0181, 0xFF),
            new OfficeArtProperty(5, 0x0183, 0xFF0000)
        }) { Bounds = new PublisherNativeRectangle(0, 0, horizontal ? 127000 : 0, horizontal ? 0 : 127000) });
        PublisherDocument document = new PublisherPageProjector(source, new PublisherQuillText(), escher,
            new PublisherParseContext(new PublisherReadOptions(), default)).Project();
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_CUSTOM_PATH_COMMAND_UNSUPPORTED");
        OfficeShape fallback = Assert.Single(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()).OfType<OfficeDrawingShape>(),
            item => item.Shape.StrokeColor == OfficeColor.FromRgb(255, 0, 0)).Shape;
        Assert.Equal(horizontal ? 10 : 0.01, fallback.Width);
        Assert.Equal(horizontal ? 0.01 : 10, fallback.Height);
        // The positive fallback canvas is not a native two-dimensional fill
        // frame. Preserve the logical zero extent and its gradient loss.
        Assert.Null(fallback.FillGradient);
        Assert.Contains(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_GRADIENT_INVALID");
    }
}
