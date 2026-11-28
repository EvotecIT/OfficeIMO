using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointSlideStructureBenchmarks {
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_SLIDE_STRUCTURE_FIXTURE");
        _deck = string.IsNullOrWhiteSpace(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        if (ReadSections() != 20) throw new InvalidOperationException("The section benchmark did not see the expected slides.");

        using var changed = PowerPointDocument.Load(new MemoryStream(GroupAndReorderShapes(), writable: false));
        if (changed.Slides.Count != 20 || changed.Slides.Any(slide =>
                slide.Shapes.Count != 2 || slide.Shapes[0] is not PowerPointGroupShape group || slide.GetGroupChildren(group).Count != 2)) {
            throw new InvalidOperationException("The grouping benchmark did not save the expected shape order.");
        }

        using var expanded = PowerPointDocument.Load(new MemoryStream(AppendSlides(), writable: false));
        if (expanded.Slides.Count != 40 || expanded.GetSections()[3].SlideIndices.Count != 25) {
            throw new InvalidOperationException("The batch benchmark did not save the expected section membership.");
        }
    }

    internal static byte[] CreateFixture() {
        using var source = PowerPointDocument.Create();
        for (int i = 0; i < 20; i++) {
            PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
            slide.AddTextBox("Quarterly results", OfficeLength.Points(20), OfficeLength.Points(15),
                OfficeLength.Points(280), OfficeLength.Points(36));
            slide.AddTextBox("Sales rose in every region", OfficeLength.Points(20), OfficeLength.Points(65),
                OfficeLength.Points(280), OfficeLength.Points(50));
            slide.AddTextBox("Revenue and margin", OfficeLength.Points(20), OfficeLength.Points(120),
                OfficeLength.Points(280), OfficeLength.Points(30));
        }

        source.AddSection("Overview", 0);
        source.AddSection("Background", 5);
        source.AddSection("Results", 10);
        source.AddSection("Appendix", 15);
        return source.ToBytes();
    }

    [Benchmark]
    public int ReadSections() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return deck.GetSections().Sum(section => section.SlideIndices.Count);
    }

    [Benchmark]
    public byte[] GroupAndReorderShapes() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in deck.Slides) {
            PowerPointGroupShape group = slide.GroupShapes(slide.Shapes.Skip(1));
            slide.SendToBack(group);
        }
        return deck.ToBytes();
    }

    [Benchmark]
    public byte[] AppendSlides() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        deck.AddSlides(20);
        return deck.ToBytes();
    }
}
