using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointTimelineBenchmarks {
    private const int SlideCount = 20;
    private const uint MotionId = 1;
    private const uint RotationId = 4;
    private const uint CommandId = 6;
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_TIMELINE_FIXTURE");
        _deck = string.IsNullOrWhiteSpace(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        using var authored = PowerPointDocument.Load(new MemoryStream(AppendActionsAndSave(), writable: false));
        if (authored.Slides.Count != SlideCount || authored.Slides.Any(slide => !slide.RemoveAnimation(CommandId))) {
            throw new InvalidOperationException("The timeline benchmark did not save each appended command.");
        }
        using var edited = PowerPointDocument.Load(new MemoryStream(EditActionsAndSave(), writable: false));
        if (edited.Slides.Count != SlideCount || edited.Slides.Any(slide =>
            slide.SetAnimationDuration(MotionId, 900) || !slide.SetAnimationDuration(RotationId, 900))) {
            throw new InvalidOperationException("The timeline benchmark did not remove and retain the intended actions.");
        }
    }

    internal static byte[] CreateFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = document.AddSlide(PowerPointSlideLayoutType.Blank);
            PowerPointTextBox shape = slide.AddTextBox("Animated title");
            PowerPointTimelineAction motion = slide.AddMotionAnimation(shape, "M 0 0 L 1 0 E");
            PowerPointTimelineAction rotation = slide.AddRotationAnimation(shape, 45);
            if (motion.TimingId != MotionId || rotation.TimingId != RotationId) {
                throw new InvalidOperationException("The timeline benchmark fixture has unexpected action IDs.");
            }
        }
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] AppendActionsAndSave() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in document.Slides) {
            PowerPointTextBox shape = slide.TextBoxes.Single();
            slide.AddCommandAnimation(shape, "play", new PowerPointTimelineOptions { Start = PowerPointTimelineStart.AfterPrevious });
            slide.AddRotationAnimation(shape, 90, new PowerPointTimelineOptions { Start = PowerPointTimelineStart.WithPrevious });
        }
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] EditActionsAndSave() {
        using var document = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in document.Slides) {
            slide.SetAnimationDuration(RotationId, 750);
            slide.RemoveAnimation(MotionId);
        }
        return document.ToBytes();
    }
}
