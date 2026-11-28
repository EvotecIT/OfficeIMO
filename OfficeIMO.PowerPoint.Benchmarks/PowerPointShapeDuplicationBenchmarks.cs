using BenchmarkDotNet.Attributes;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointShapeDuplicationBenchmarks {
    private const int SlideCount = 24;
    private static readonly OfficeLength Shift = OfficeLength.Points(12);
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_DUPLICATION_FIXTURE");
        _deck = string.IsNullOrWhiteSpace(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        using var singles = PowerPointDocument.Load(new MemoryStream(DuplicateShapesAndSave(), writable: false));
        if (singles.Slides.Count != SlideCount || singles.Slides.Any(slide => {
            PowerPointAutoShape[] boxes = slide.Shapes.OfType<PowerPointAutoShape>().ToArray();
            return boxes.Length != 2 || boxes[1].Top - boxes[0].Top != Shift;
        })) {
            throw new InvalidOperationException("The benchmark did not save each duplicated shape.");
        }
        using var groups = PowerPointDocument.Load(new MemoryStream(DuplicateGroupsAndSave(), writable: false));
        if (groups.Slides.Count != SlideCount || groups.Slides.Any(slide => {
            PowerPointGroupShape[] copied = slide.Shapes.OfType<PowerPointGroupShape>().ToArray();
            return copied.Length != 2 || copied[1].Left - copied[0].Left != Shift || copied[0].Id == copied[1].Id;
        })) {
            throw new InvalidOperationException("The benchmark did not save each duplicated group.");
        }
    }

    internal static byte[] CreateFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = document.AddSlide(PowerPointSlideLayoutType.Blank);
            slide.AddRectangle(OfficeLength.Points(10), OfficeLength.Points(10), OfficeLength.Points(70), OfficeLength.Points(30), "Standalone");
            PowerPointAutoShape ellipse = slide.AddShape(OfficePresetShapeType.Ellipse, OfficeLength.Points(95), OfficeLength.Points(100));
            PowerPointTextBox text = slide.AddTextBox("Grouped text", left: OfficeLength.Points(245), top: OfficeLength.Points(125));
            slide.GroupShapes(new PowerPointShape[] { ellipse, text });
        }
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] DuplicateShapesAndSave() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in document.Slides) {
            slide.DuplicateShape(slide.Shapes.OfType<PowerPointAutoShape>().Single(), Shift, Shift);
        }
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] DuplicateGroupsAndSave() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in document.Slides) {
            slide.DuplicateShape(slide.Shapes.OfType<PowerPointGroupShape>().Single(), Shift, Shift);
        }
        return document.ToBytes();
    }
}
