using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointClassicCommentReassignmentBenchmarks {
    private byte[] _fixture = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixturePath = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLASSIC_COMMENTS_FIXTURE");
        _fixture = fixturePath == null ? PowerPointClassicCommentReadBenchmarks.CreateFixture() : File.ReadAllBytes(fixturePath);
        Validate(ReassignToExistingAuthor(), "Presentation Reviewer", 13U);
        Validate(ReassignToNewAuthor(), "Additional Reviewer", 1U);
    }

    [Benchmark]
    public byte[] ReassignToExistingAuthor() => ReassignAndSave(new PowerPointCommentAuthor("Presentation Reviewer", "PR"));

    [Benchmark]
    public byte[] ReassignToNewAuthor() => ReassignAndSave(new PowerPointCommentAuthor("Additional Reviewer", "AR"));

    private byte[] ReassignAndSave(PowerPointCommentAuthor author) {
        using var deck = PowerPointDocument.Load(new MemoryStream(_fixture, writable: false));
        deck.GetClassicComments(deck.Slides[11])[0].SetAuthor(author);
        return deck.ToBytes();
    }

    private static void Validate(byte[] bytes, string name, uint expectedIndex) {
        using var deck = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        IReadOnlyList<PowerPointClassicComment> notes = deck.GetClassicComments(deck.Slides[11]);
        if (notes.Count != 2 || notes[0].Author.Name != name || notes[0].Index != expectedIndex ||
            notes[0].Text != "An editorial note" || notes[1].Text != "A reviewer response") {
            throw new InvalidOperationException("Classic comment reassignment did not round trip.");
        }
    }
}
