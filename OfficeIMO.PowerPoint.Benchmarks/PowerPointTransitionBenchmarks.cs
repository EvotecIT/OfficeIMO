using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointTransitionBenchmarks {
    private const int SlideCount = 20;
    private static readonly byte[] Audio = { 82, 73, 70, 70, 4, 0, 0, 0, 87, 65, 86, 69 };
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_TRANSITION_FIXTURE");
        _deck = string.IsNullOrWhiteSpace(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        if (ReadSlideTransitions() != SlideCount) {
            throw new InvalidOperationException("The transition benchmark did not read the expected effects.");
        }
        using var edited = PowerPointDocument.Load(new MemoryStream(EditAndSaveTransitions(), writable: false));
        if (edited.Slides.Count != SlideCount ||
            edited.Slides.Any(slide => slide.Transition != PowerPointSlideTransition.WipeDown ||
                slide.TransitionDurationSeconds != 0.75 || slide.TransitionAdvanceOnClick != false)) {
            throw new InvalidOperationException("The transition benchmark did not save the expected effects.");
        }
        using var sounds = PowerPointDocument.Load(new MemoryStream(EmbedAndSaveTransitionSounds(), writable: false));
        if (sounds.Slides.Count != SlideCount ||
            sounds.Slides.Any(slide => slide.GetTransitionSoundBytes()?.SequenceEqual(Audio) != true)) {
            throw new InvalidOperationException("The transition benchmark did not save the expected sounds.");
        }
    }

    internal static byte[] CreateFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = document.AddSlide(PowerPointSlideLayoutType.Blank);
            slide.AddTextBox("Quarterly review");
            slide.Transition = index % 2 == 0 ? PowerPointSlideTransition.Fade : PowerPointSlideTransition.Morph;
            slide.TransitionSpeed = PowerPointSlideTransitionSpeed.Fast;
            slide.TransitionDurationSeconds = 0.8;
            slide.TransitionAdvanceAfterSeconds = 2.5;
        }

        return document.ToBytes();
    }

    [Benchmark]
    public int ReadSlideTransitions() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return document.Slides.Count(slide =>
            slide.Transition is PowerPointSlideTransition.Fade or PowerPointSlideTransition.Morph &&
            slide.TransitionSpeed == PowerPointSlideTransitionSpeed.Fast &&
            slide.TransitionDurationSeconds == 0.8 && slide.TransitionAdvanceAfterSeconds == 2.5);
    }

    [Benchmark]
    public byte[] EditAndSaveTransitions() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in document.Slides) {
            slide.Transition = PowerPointSlideTransition.WipeDown;
            slide.TransitionDurationSeconds = 0.75;
            slide.TransitionAdvanceOnClick = false;
        }

        return document.ToBytes();
    }

    [Benchmark]
    public byte[] EmbedAndSaveTransitionSounds() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        using var sound = new MemoryStream(Audio, writable: false);
        foreach (PowerPointSlide slide in document.Slides) {
            slide.SetTransitionSound(sound, "Sound", loop: true);
        }
        return document.ToBytes();
    }
}
