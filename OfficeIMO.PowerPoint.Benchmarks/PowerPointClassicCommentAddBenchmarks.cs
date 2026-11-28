using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointClassicCommentAddBenchmarks {
    private byte[] _empty = Array.Empty<byte>();
    private byte[] _populated = Array.Empty<byte>();
    private static readonly DateTime Date = new DateTime(2026, 4, 20, 9, 0, 0, DateTimeKind.Utc);

    [GlobalSetup]
    public void Setup() {
        string? emptyFixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLASSIC_COMMENT_ADD_EMPTY");
        string? populatedFixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLASSIC_COMMENT_ADD_POPULATED");
        _empty = emptyFixture == null ? CreateEmptyFixture() : File.ReadAllBytes(emptyFixture);
        _populated = populatedFixture == null ? PowerPointClassicCommentReadBenchmarks.CreateFixture() : File.ReadAllBytes(populatedFixture);
        Validate(AddFirstComment(), 1, "Presentation Editor");
        Validate(AppendToExistingComments(), 3, "Presentation Editor");
        Validate(AddAcrossDeck(), 3, "Presentation Editor", modifiedSlides: 12);
    }

    [Benchmark]
    public byte[] AddFirstComment() => AddAndSave(_empty);

    [Benchmark]
    public byte[] AppendToExistingComments() => AddAndSave(_populated);

    [Benchmark]
    public byte[] AddAcrossDeck() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_populated, writable: false));
        var author = new PowerPointCommentAuthor("Presentation Editor", "PE");
        for (int index = 0; index < 12; index++) {
            deck.AddClassicComment(deck.Slides[index], author, "A later review note",
                OfficeLength.Points(46), OfficeLength.Points(-28), Date);
        }
        return deck.ToBytes();
    }

    private static byte[] AddAndSave(byte[] bytes) {
        using var deck = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        deck.AddClassicComment(deck.Slides[0], new PowerPointCommentAuthor("Presentation Editor", "PE"),
            "A later review note", OfficeLength.Points(46), OfficeLength.Points(-28), Date);
        return deck.ToBytes();
    }

    private static void Validate(byte[] bytes, int expectedCount, string author, int modifiedSlides = 1) {
        using var document = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        for (int index = 0; index < modifiedSlides; index++) {
            IReadOnlyList<PowerPointClassicComment> comments = document.GetClassicComments(document.Slides[index]);
            if (comments.Count != expectedCount || comments[comments.Count - 1].Text != "A later review note" ||
                comments[comments.Count - 1].Author.Name != author || comments[comments.Count - 1].Created != Date) {
                throw new InvalidOperationException("The added classic comment did not round trip.");
            }
        }
    }

    internal static byte[] CreateEmptyFixture() {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < 24; index++) document.AddSlide().AddTextBox("Review slide " + (index + 1));
        return document.ToBytes();
    }
}
