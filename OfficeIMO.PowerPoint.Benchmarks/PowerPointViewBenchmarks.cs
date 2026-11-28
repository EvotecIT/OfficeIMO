using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointViewBenchmarks {
    private const int GuideCount = 32;
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_VIEW_FIXTURE");
        _deck = string.IsNullOrWhiteSpace(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        if (ReadGridAndGuides() != (true, 19050L, 38100L, GuideCount)) {
            throw new InvalidOperationException("The view benchmark did not read the expected grid and guides.");
        }
        using var edited = PowerPointDocument.Load(new MemoryStream(EditViewAndSave(), writable: false));
        if (edited.SnapToGrid || edited.GridSpacingX.ToEmus() != 63500 || edited.GetGuides().Count != GuideCount / 2) {
            throw new InvalidOperationException("The view benchmark did not save the expected settings.");
        }
    }

    internal static byte[] CreateFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < 20; index++) document.AddSlide().AddTextBox("Planning");
        document.SnapToGrid = true;
        document.SetGridSpacing(OfficeLength.Points(1.5), OfficeLength.Points(3));
        document.SetGuides(CreateGuides(GuideCount));
        return document.ToBytes();
    }

    private static IEnumerable<PowerPointGuideInfo> CreateGuides(int count) =>
        Enumerable.Range(0, count).Select(index => new PowerPointGuideInfo(
            index % 2 == 0 ? PowerPointGuideOrientation.Horizontal : PowerPointGuideOrientation.Vertical,
            OfficeLength.Points(index * 5)));

    [Benchmark]
    public (bool SnapToGrid, long GridX, long GridY, int Guides) ReadGridAndGuides() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return (document.SnapToGrid, document.GridSpacingX.ToEmus(), document.GridSpacingY.ToEmus(), document.GetGuides().Count);
    }

    [Benchmark]
    public byte[] EditViewAndSave() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        document.SnapToGrid = false;
        document.SetGridSpacing(OfficeLength.Points(5), OfficeLength.Points(4));
        document.SetGuides(CreateGuides(GuideCount / 2));
        return document.ToBytes();
    }
}
