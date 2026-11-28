using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointCustomShowBenchmarks {
    private const int Slides = 20;
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CUSTOM_SHOW_FIXTURE");
        _deck = string.IsNullOrEmpty(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        if (ReadCustomShows() != (3, 34, 9U)) {
            throw new InvalidOperationException("The fixture did not preserve all custom-show entries.");
        }

        using var saved = PowerPointDocument.Load(new MemoryStream(EditCustomShowsAndSave(), writable: false));
        IReadOnlyList<PowerPointCustomShow> shows = saved.CustomShows;
        if (shows.Count != 4 || shows[0].Name != "Executive" || shows[0].Slides.Count != 12 ||
            !ReferenceEquals(saved.Slides[8], shows[0].Slides[0]) ||
            !ReferenceEquals(saved.Slides[0], shows[0].Slides[5]) ||
            shows[3].Name != "Appendix" || !shows[3].Slides.SequenceEqual(new[] { saved.Slides[19], saved.Slides[0] })) {
            throw new InvalidOperationException("The edited deck did not preserve the requested shows and slide order.");
        }
    }

    internal static byte[] CreateFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < Slides; index++) {
            document.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("Planning slide " + index);
        }
        document.AddCustomShow("Overview", document.Slides.Take(10));
        document.AddCustomShow("Working group", document.Slides.Reverse().Take(12));
        document.AddCustomShow("Audit", document.Slides.Skip(5).Take(12));
        return document.ToBytes();
    }

    [Benchmark]
    public (int Shows, int Entries, uint Lookup) ReadCustomShows() {
        using PowerPointDocument document = Load();
        IReadOnlyList<PowerPointCustomShow> shows = document.CustomShows;
        int entries = 0;
        uint ids = document.GetCustomShow(" audit ")!.Id;
        foreach (PowerPointCustomShow show in shows) {
            entries += show.Slides.Count;
            ids += show.Id;
        }
        return (shows.Count, entries, ids);
    }

    [Benchmark]
    public byte[] EditCustomShowsAndSave() {
        using PowerPointDocument document = Load();
        PowerPointCustomShow show = document.GetCustomShow("Overview")!;
        document.RenameCustomShow(show, "Executive");
        show.MoveSlide(8, 0).InsertSlide(5, document.Slides[0]).AddSlide(document.Slides[14]);
        document.AddCustomShow("Appendix", new[] { document.Slides[19], document.Slides[0] });
        return document.ToBytes();
    }

    private PowerPointDocument Load() => PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
}
