using BenchmarkDotNet.Attributes;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointConnectorGeometryBenchmarks {
    private const int SlideCount = 20;
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_GEOMETRY_FIXTURE");
        _deck = string.IsNullOrWhiteSpace(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        if (ReadConnectorGeometry() != SlideCount) {
            throw new InvalidOperationException("The connector benchmark could not resolve all its glued ends.");
        }
        using var authored = PowerPointDocument.Load(new MemoryStream(AddAndSaveShapes(), writable: false));
        if (authored.Slides.Count != SlideCount || authored.Slides.Any(slide =>
                slide.Shapes.Count != 8 || slide.Shapes.OfType<PowerPointConnectionShape>().Count() != 3)) {
            throw new InvalidOperationException("The authoring benchmark did not save the expected shapes and connectors.");
        }
        using var rerouted = PowerPointDocument.Load(new MemoryStream(RerouteAndSaveConnectors(), writable: false));
        if (rerouted.Slides.Count != SlideCount || rerouted.Slides.Any(slide => {
                var connector = slide.Shapes.OfType<PowerPointConnectionShape>().Single();
                return connector.EndShape?.Top != OfficeLength.Points(175) || connector.Height.ToEmus() == 0;
            })) {
            throw new InvalidOperationException("The edit benchmark did not reroute all its connectors.");
        }
    }

    internal static byte[] CreateFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = document.AddSlide(PowerPointSlideLayoutType.Blank);
            PowerPointAutoShape left = AddBox(slide, OfficePresetShapeType.Rectangle, "Left", 30, 45);
            PowerPointAutoShape right = AddBox(slide, OfficePresetShapeType.Ellipse, "Right", 340, 45);
            slide.AddConnector(left, right, PowerPointConnectorKind.Straight);
        }
        return document.ToBytes();
    }

    internal static void AddShapes(PowerPointSlide slide) {
        PowerPointAutoShape left = AddBox(slide, OfficePresetShapeType.Triangle, "New left", 55, 210);
        PowerPointAutoShape right = AddBox(slide, OfficePresetShapeType.RoundRectangle, "New right", 360, 250);
        OfficeShape geometry = OfficeShape.Path(OfficePathCommand.MoveTo(0, 0),
            OfficePathCommand.QuadraticBezierTo(45, 0, 50, 25), OfficePathCommand.CubicBezierTo(70, 45, 82, 100, 100, 50),
            OfficePathCommand.LineTo(100, 100), OfficePathCommand.Close());
        geometry.FillRule = OfficeFillRule.NonZero;
        geometry.FillColor = OfficeColor.FromRgb(14, 165, 233);
        geometry.StrokeColor = OfficeColor.FromRgb(12, 74, 110);
        geometry.StrokeWidth = 2.25;
        slide.AddCustomGeometry(geometry, OfficeLength.Points(130), OfficeLength.Points(190), OfficeLength.Points(90),
            OfficeLength.Points(70), "Freeform");
        slide.AddConnector(left, right, PowerPointConnectorKind.Elbow);
        slide.AddConnector(OfficeLength.Points(450), OfficeLength.Points(300), OfficeLength.Points(220), OfficeLength.Points(130),
            PowerPointConnectorKind.Curved);
    }

    private static PowerPointAutoShape AddBox(PowerPointSlide slide, OfficePresetShapeType type, string name, int x, int y) =>
        slide.AddShape(type, OfficeLength.Points(x), OfficeLength.Points(y), OfficeLength.Points(85), OfficeLength.Points(50), name);

    [Benchmark]
    public int ReadConnectorGeometry() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return document.Slides.SelectMany(slide => slide.Shapes.OfType<PowerPointConnectionShape>())
            .Count(connector => connector.Kind == PowerPointConnectorKind.Straight &&
                connector.StartShape?.Name == "Left" && connector.EndShape?.Name == "Right");
    }

    [Benchmark]
    public byte[] AddAndSaveShapes() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in document.Slides) AddShapes(slide);
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] RerouteAndSaveConnectors() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in document.Slides) slide.GetShapeByName("Right")!.Top = OfficeLength.Points(175);
        return document.ToBytes();
    }
}
