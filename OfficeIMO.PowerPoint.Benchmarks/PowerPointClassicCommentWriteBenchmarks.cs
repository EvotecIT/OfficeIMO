using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointClassicCommentWriteBenchmarks {
    private static readonly DateTime RevisedDate = new DateTime(2026, 4, 20, 9, 0, 0, DateTimeKind.Utc);
    private byte[] _source = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLASSIC_COMMENT_WRITE_FIXTURE");
        _source = fixture == null ? PowerPointClassicCommentReadBenchmarks.CreateFixture() : File.ReadAllBytes(fixture);
        using var editedText = PowerPointDocument.Load(new MemoryStream(EditText(), writable: false));
        PowerPointClassicComment text = editedText.GetClassicComments(editedText.Slides[0])[0];
        if (text.Text != "Prepared review text\nSecond line.") {
            throw new InvalidOperationException("The comment text was not saved.");
        }
        using var editedFields = PowerPointDocument.Load(new MemoryStream(EditPositionAndDate(), writable: false));
        PowerPointClassicComment fields = editedFields.GetClassicComments(editedFields.Slides[0])[0];
        if (fields.X != OfficeLength.Points(-42) || fields.Y != OfficeLength.Points(135) || fields.Created != RevisedDate) {
            throw new InvalidOperationException("The comment position or timestamp was not saved.");
        }
    }

    [Benchmark]
    public byte[] EditText() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_source, writable: false));
        PowerPointClassicComment comment = deck.GetClassicComments(deck.Slides[0])[0];
        comment.Text = "Prepared review text\nSecond line.";
        return deck.ToBytes();
    }

    [Benchmark]
    public byte[] EditPositionAndDate() {
        using var deck = PowerPointDocument.Load(new MemoryStream(_source, writable: false));
        PowerPointClassicComment comment = deck.GetClassicComments(deck.Slides[0])[0];
        comment.X = OfficeLength.Points(-42);
        comment.Y = OfficeLength.Points(135);
        comment.Created = RevisedDate;
        return deck.ToBytes();
    }
}
