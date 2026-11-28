using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointClassicCommentSlideRemovalBenchmarks {
    private byte[] _fixture = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixturePath = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLASSIC_COMMENTS_FIXTURE");
        _fixture = fixturePath == null ? PowerPointClassicCommentReadBenchmarks.CreateFixture() : File.ReadAllBytes(fixturePath);
        Validate(RemoveCommentedSlide(), 23, 22);
        Validate(RemoveAllCommentedSlides(), 12, 0);
    }

    [Benchmark]
    public byte[] RemoveCommentedSlide() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_fixture, writable: false));
        deck.RemoveSlide(11);
        return deck.ToBytes();
    }

    [Benchmark]
    public byte[] RemoveAllCommentedSlides() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_fixture, writable: false));
        for (int index = 11; index >= 0; index--) deck.RemoveSlide(index);
        return deck.ToBytes();
    }

    private static void Validate(byte[] bytes, int expectedSlides, int expectedComments) {
        using var deck = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        if (deck.Slides.Count != expectedSlides || deck.Slides.Sum(slide => deck.GetClassicComments(slide).Count) != expectedComments) {
            throw new InvalidOperationException("Removing commented slides did not round trip.");
        }
    }
}
