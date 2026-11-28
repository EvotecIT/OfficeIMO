using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointClassicAnimationWriteBenchmarks {
    private const int SlideCount = 20;
    private static readonly PowerPointClassicAnimationOptions UpdatedOptions = new() {
        Automatic = true, DelayMilliseconds = 300, AfterEffect = PowerPointClassicAnimationAfterEffect.Dim,
        RawDimColor = 0xFF6080A0U, StopsSound = true
    };
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLASSIC_WRITE_FIXTURE");
        _deck = string.IsNullOrEmpty(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        ValidateSavedAnimations(AuthorClassicAnimationsAndSave(), 0);
        ValidateSavedAnimations(ReplaceClassicAnimationsAndSave(), 1);
    }

    private static void ValidateSavedAnimations(byte[] bytes, int firstSlide) {
        using var presentation = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        for (int index = firstSlide; index < presentation.Slides.Count; index += 2) {
            PowerPointClassicAnimation animation = presentation.Slides[index].ClassicAnimations.Single();
            if (animation.DelayMilliseconds != UpdatedOptions.DelayMilliseconds || !animation.StopsSound
                || animation.Effect != PowerPointClassicAnimationEffect.Fade || animation.AfterEffect != UpdatedOptions.AfterEffect) {
                throw new InvalidOperationException("The classic timing change was not saved.");
            }
        }
    }

    internal static byte[] CreateFixture() {
        using var presentation = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
            PowerPointTextBox target = slide.AddTextBox("Classic animation");
            if (index % 2 == 1) slide.AddClassicAnimation(target, PowerPointClassicAnimationEffect.Wipe);
        }
        return presentation.ToBytes();
    }

    [Benchmark]
    public byte[] AuthorClassicAnimationsAndSave() => SetAnimationsAndSave(0);

    [Benchmark]
    public byte[] ReplaceClassicAnimationsAndSave() => SetAnimationsAndSave(1);

    private byte[] SetAnimationsAndSave(int firstSlide) {
        using var presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        for (int index = firstSlide; index < presentation.Slides.Count; index += 2) {
            PowerPointSlide slide = presentation.Slides[index];
            slide.AddClassicAnimation(slide.TextBoxes.Single(), PowerPointClassicAnimationEffect.Fade, UpdatedOptions);
        }
        return presentation.ToBytes();
    }
}
