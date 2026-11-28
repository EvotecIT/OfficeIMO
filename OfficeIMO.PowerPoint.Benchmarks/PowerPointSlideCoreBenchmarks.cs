using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointSlideCoreBenchmarks {
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_SLIDE_CORE_FIXTURE");
        _deck = string.IsNullOrWhiteSpace(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        if (ReadSlideScenes() != 60) throw new InvalidOperationException("The scene benchmark did not see the expected text spans.");
        using var edited = PowerPointDocument.Load(new MemoryStream(EditSlideMetadataAndBackgrounds(), writable: false));
        if (edited.Slides.Count != 20 || edited.Slides.Any(slide =>
                slide.Name?.StartsWith("Revised ", StringComparison.Ordinal) != true ||
                slide.GetBackground().Kind != PowerPointSlideBackgroundKind.LinearGradient)
            || edited.SlideSize.Width != OfficeLength.Points(336)) {
            throw new InvalidOperationException("The edit benchmark did not save the expected presentation.");
        }
    }

    internal static byte[] CreateFixture() {
        using var source = PowerPointDocument.Create();
        source.SlideSize.SetSize(OfficeLength.Points(320), OfficeLength.Points(180));
        PowerPointSlideMaster master = source.SlideMasters[0];
        master.BackgroundColor = OfficeColor.Parse("E7EDF4");
        master.AddTextBox("Master footer", OfficeLength.Points(20), OfficeLength.Points(153),
            OfficeLength.Points(200), OfficeLength.Points(20));

        for (int index = 0; index < 20; index++) {
            PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
            slide.Name = "Slide " + index;
            slide.AddTextBox("Quarterly performance", OfficeLength.Points(20), OfficeLength.Points(15),
                OfficeLength.Points(280), OfficeLength.Points(36));
            slide.AddTextBox("Revenue grew in every region", OfficeLength.Points(20), OfficeLength.Points(72),
                OfficeLength.Points(280), OfficeLength.Points(50));
        }

        return source.ToBytes();
    }

    [Benchmark]
    public int ReadSlideScenes() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        int textSpans = 0;
        foreach (PowerPointSlide slide in deck.Slides) {
            textSpans += slide.CreateVisualSnapshot().TextSpans.Count;
        }

        return textSpans;
    }

    [Benchmark]
    public byte[] EditSlideMetadataAndBackgrounds() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        deck.SlideSize.Width = OfficeLength.Points(336);
        foreach (PowerPointSlide slide in deck.Slides) {
            slide.Name = "Revised " + slide.Name;
            slide.SetBackgroundGradient(OfficeColor.Parse("E7EDF4"), OfficeColor.Parse("FFFFFF"), 45D);
        }

        return deck.ToBytes();
    }
}
