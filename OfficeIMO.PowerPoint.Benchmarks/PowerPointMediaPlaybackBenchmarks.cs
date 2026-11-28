using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointMediaPlaybackBenchmarks {
    private const int SlideCount = 20;
    private static readonly Uri OriginalSource = new("https://example.test/media/original.mp4");
    private static readonly Uri NewSource = new("https://example.test/media/updated%20clip.mp4");
    private static readonly PowerPointMediaPlaybackOptions UpdatedOptions = new() {
        VolumePercent = 55, Mute = true, Loop = true, FullScreen = true, ShowWhenStopped = false,
        TrimStartMilliseconds = 150, TrimEndMilliseconds = 2550, FadeInMilliseconds = 300, FadeOutMilliseconds = 200
    };
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_MEDIA_PLAYBACK_FIXTURE");
        _deck = string.IsNullOrEmpty(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        if (ReadPlayback() != SlideCount * 80) throw new InvalidOperationException("The fixture has unexpected playback metadata.");
        using var updated = PowerPointDocument.Load(new MemoryStream(EditPlaybackAndSave(), writable: false));
        if (updated.Slides.Count != SlideCount || updated.Slides.Any(slide => !IsUpdated(slide.Media.Single()))) {
            throw new InvalidOperationException("The media playback benchmark did not save the new playback metadata.");
        }
    }

    private static bool IsUpdated(PowerPointMedia media) {
        PowerPointMediaPlaybackOptions options = media.GetPlaybackOptions();
        return media.LinkUri == NewSource && options.VolumePercent == UpdatedOptions.VolumePercent && options.Mute && options.FullScreen
            && options.Loop && !options.ShowWhenStopped && options.TrimStartMilliseconds == UpdatedOptions.TrimStartMilliseconds
            && options.TrimEndMilliseconds == UpdatedOptions.TrimEndMilliseconds && options.FadeInMilliseconds == UpdatedOptions.FadeInMilliseconds
            && options.FadeOutMilliseconds == UpdatedOptions.FadeOutMilliseconds;
    }

    internal static byte[] CreateFixture() {
        using var deck = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = deck.AddSlide(PowerPointSlideLayoutType.Blank);
            slide.AddLinkedVideo(OriginalSource);
        }
        return deck.ToBytes();
    }

    [Benchmark]
    public int ReadPlayback() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return deck.Slides.Sum(slide => slide.Media.Single().GetPlaybackOptions().VolumePercent);
    }

    [Benchmark]
    public byte[] EditPlaybackAndSave() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in deck.Slides) {
            PowerPointMedia media = slide.Media.Single();
            media.UpdateLink(NewSource);
            media.SetPlaybackOptions(UpdatedOptions);
        }
        return deck.ToBytes();
    }
}
