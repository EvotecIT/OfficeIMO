using BenchmarkDotNet.Attributes;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointMediaBenchmarks {
    private const int SlideCount = 12;
    private static readonly Uri VideoLink = new("https://example.test/media/video.mp4");
    private static readonly byte[] Poster = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==");
    private static readonly byte[] Video = Enumerable.Range(0, 256 * 1024).Select(offset => (byte)(offset ^ (offset >> 8))).ToArray();
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_MEDIA_FIXTURE");
        _deck = string.IsNullOrEmpty(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        if (ReadEmbeddedVideo() != SlideCount * Video.Length) {
            throw new InvalidOperationException("The media benchmark input did not contain all video bytes.");
        }
        using var saved = PowerPointDocument.Load(new MemoryStream(AddVideosAndSave(), writable: false));
        if (saved.Slides.Count != SlideCount || saved.Slides.Any(slide => slide.Media.Count() != 3
            || !slide.Media.Take(2).All(item => item.GetData().SequenceEqual(Video))
            || slide.Media.Last().LinkUri != VideoLink)) {
            throw new InvalidOperationException("The media benchmark did not retain the authored payload and link.");
        }
    }

    internal static byte[] CreateFixture() {
        using var deck = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = deck.AddSlide(PowerPointSlideLayoutType.Blank);
            slide.AddTextBox("Media");
            AddVideo(slide);
        }
        return deck.ToBytes();
    }

    internal static PowerPointMedia AddVideo(PowerPointSlide slide) {
        using var data = new MemoryStream(Video, writable: false);
        using var poster = new MemoryStream(Poster, writable: false);
        return slide.AddVideo(data, "video/mp4", ".mp4", posterImage: poster, posterImageType: OfficeImageFormat.Png);
    }

    [Benchmark]
    public int ReadEmbeddedVideo() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return deck.Slides.Sum(slide => slide.Media.Single().GetData().Length);
    }

    [Benchmark]
    public byte[] AddVideosAndSave() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in deck.Slides) {
            AddVideo(slide);
            slide.AddLinkedVideo(VideoLink);
        }
        return deck.ToBytes();
    }
}
