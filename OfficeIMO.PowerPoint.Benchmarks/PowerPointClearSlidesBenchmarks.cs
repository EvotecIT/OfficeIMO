using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointClearSlidesBenchmarks {
    internal const string NewText = "New slide in the cleared deck";
    private byte[] _fixture = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? path = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLEAR_SLIDES_FIXTURE");
        _fixture = path == null ? CreateFixture() : File.ReadAllBytes(path);
        using var cleared = Load(ClearAndSave());
        using var reused = Load(ClearAndReuse());
        if (cleared.Slides.Count != 0 || reused.Slides.Count != 1 || reused.Slides[0].TextBoxes.Single().Text != NewText) {
            throw new InvalidOperationException("The cleared deck did not retain a usable template.");
        }
    }

    [Benchmark]
    public byte[] ClearAndSave() => Save(reuse: false);

    [Benchmark]
    public byte[] ClearAndReuse() => Save(reuse: true);

    private byte[] Save(bool reuse) {
        using var document = Load(_fixture);
        document.ClearSlides();
        if (reuse) document.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox(NewText);
        return document.ToBytes();
    }

    private static PowerPointDocument Load(byte[] data) => PowerPointDocument.Load(new MemoryStream(data, writable: false));

    internal static byte[] CreateFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < 48; index++) {
            PowerPointSlide slide = document.AddSlide(PowerPointSlideLayoutType.Blank);
            slide.AddTextBox("Slide retained in the source only");
        }
        document.AddSection("First section", 0);
        document.AddSection("Second section", 24);
        document.AddCustomShow("Show in the source", new[] { document.Slides[0], document.Slides[24] });
        return document.ToBytes();
    }
}
