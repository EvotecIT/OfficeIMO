using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointClassicCommentRemovalBenchmarks {
    private byte[] _fixture = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixturePath = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLASSIC_COMMENTS_FIXTURE");
        _fixture = fixturePath == null ? PowerPointClassicCommentReadBenchmarks.CreateFixture() : File.ReadAllBytes(fixturePath);
        Validate(RemoveCommentWithSurvivor(), 1);
        Validate(RemoveLastCommentOnSlide(), 0);
        using var emptied = PowerPointDocument.Load(new MemoryStream(RemoveAcrossDeck(), writable: false));
        if (emptied.Slides.Any(slide => emptied.GetClassicComments(slide).Count != 0)) {
            throw new InvalidOperationException("The deck still has classic comments.");
        }
    }

    [Benchmark]
    public byte[] RemoveCommentWithSurvivor() => RemoveAndSave(removeAllOnSlide: false);

    [Benchmark]
    public byte[] RemoveLastCommentOnSlide() => RemoveAndSave(removeAllOnSlide: true);

    [Benchmark]
    public byte[] RemoveAcrossDeck() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_fixture, writable: false));
        for (int index = 0; index < 12; index++) {
            foreach (PowerPointClassicComment comment in deck.GetClassicComments(deck.Slides[index])) comment.Remove();
        }
        return deck.ToBytes();
    }

    private byte[] RemoveAndSave(bool removeAllOnSlide) {
        using var deck = PowerPointDocument.Load(new MemoryStream(_fixture, writable: false));
        IReadOnlyList<PowerPointClassicComment> comments = deck.GetClassicComments(deck.Slides[11]);
        comments[0].Remove();
        if (removeAllOnSlide) comments[1].Remove();
        return deck.ToBytes();
    }

    private static void Validate(byte[] bytes, int expectedCount) {
        using var deck = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        IReadOnlyList<PowerPointClassicComment> remaining = deck.GetClassicComments(deck.Slides[11]);
        if (remaining.Count != expectedCount || expectedCount > 0 && remaining[0].Text != "A reviewer response" ||
            deck.GetClassicComments(deck.Slides[10]).Count != 2) {
            throw new InvalidOperationException("Classic comment removal did not round trip.");
        }
    }
}
